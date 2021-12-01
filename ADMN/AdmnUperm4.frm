VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnUperm4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Permissions"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGrt 
      Caption         =   "&All"
      Height          =   360
      Left            =   3120
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Grant All"
      Top             =   0
      Width           =   720
   End
   Begin VB.CommandButton cmdRvk 
      Caption         =   "&None"
      Height          =   360
      Left            =   3960
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Revoke All"
      Top             =   0
      Width           =   720
   End
   Begin VB.CheckBox optGr1 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox optGr2 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox optGr3 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox optGr4 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   20
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox optGr5 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   44
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox optGr6 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   27
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optGr7 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1860
      TabIndex        =   29
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optGr8 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1680
      TabIndex        =   35
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optFn8 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4440
      TabIndex        =   39
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optVw8 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3480
      TabIndex        =   38
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optEd8 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2520
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optFn7 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4620
      TabIndex        =   34
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optVw7 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3660
      TabIndex        =   33
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optEd7 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2700
      TabIndex        =   32
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optFn6 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4620
      TabIndex        =   28
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optVw6 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3660
      TabIndex        =   31
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optEd6 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2700
      TabIndex        =   30
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox optFn5 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4620
      TabIndex        =   26
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox optVw5 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3660
      TabIndex        =   25
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox optEd5 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2700
      TabIndex        =   24
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox optFn4 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4620
      TabIndex        =   23
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox optVw4 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3660
      TabIndex        =   22
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox optEd4 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2700
      TabIndex        =   21
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox optFn3 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4620
      TabIndex        =   19
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox optVw3 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3660
      TabIndex        =   18
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox optEd3 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2700
      TabIndex        =   17
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox optFn2 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4620
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox optVw2 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3660
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox optEd2 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2700
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox optFn1 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4620
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox optVw1 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3660
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox optEd1 
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2700
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   720
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2490
      FormDesignWidth =   5745
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   1860
      TabIndex        =   45
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblSect 
      BackStyle       =   0  'Transparent
      Caption         =   "Production Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Functions   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   4620
      TabIndex        =   42
      Top             =   480
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "View         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3660
      TabIndex        =   41
      Top             =   480
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2700
      TabIndex        =   40
      Top             =   480
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Eight"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   36
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Seven"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label V 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor RFQ's"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MRP"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PAC"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasing"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "CRP"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO's"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "AdmnUperm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions

Dim bOnLoad As Byte

Private Sub FormatBoxes()
   '    Dim bControl As Byte
   '        For bControl = 0 To Controls.Count - 1
   '            If TypeOf Controls(bControl) Is CheckBox Then _
   '                Controls(bControl).Caption = "____"
   '        Next
   
End Sub


Private Sub TestBoxes(sSwitch As Byte, bValue As Byte)
   Select Case sSwitch
      Case 1
         optEd1.Enabled = bValue
         optVw1.Enabled = bValue
         optFn1.Enabled = bValue
      Case 2
         optEd2.Enabled = bValue
         optVw2.Enabled = bValue
         optFn2.Enabled = bValue
      Case 3
         optEd3.Enabled = bValue
         optVw3.Enabled = bValue
         optFn3.Enabled = bValue
      Case 4
         optEd4.Enabled = bValue
         optVw4.Enabled = bValue
         optFn4.Enabled = bValue
      Case 5
         optEd5.Enabled = bValue
         optVw5.Enabled = bValue
         optFn5.Enabled = bValue
      Case 6
         optEd6.Enabled = bValue
         optVw6.Enabled = bValue
         optFn6.Enabled = bValue
      Case 7
         optEd7.Enabled = bValue
         optVw7.Enabled = bValue
         optFn7.Enabled = bValue
      Case Else
         optEd8.Enabled = bValue
         optVw8.Enabled = bValue
         optFn8.Enabled = bValue
   End Select
   
End Sub

Private Sub CheckSection()
   'Check Section-Must be changed as added or removed
   Dim b As Byte
   'Deny new sections
   If Trim(AdmnUuser2.cmbGrp) = "Users" Then
      optGr7.Value = vbUnchecked
      optGr8.Value = vbUnchecked
   End If
   If optGr1.Value = 0 And optGr2.Value = 0 And optGr3.Value = 0 _
                       And optGr4.Value = 0 And optGr5.Value = 0 And optGr6.Value = 0 _
                       And optGr7.Value = 0 And optGr8.Value = 0 Then
      b = 0
   Else
      b = 1
   End If
   Secure.UserProd = b
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub cmdCan_Click()
   CheckSection
   Unload Me
   
End Sub



Private Sub cmdGrt_Click()
   GrantAll
   
End Sub

Private Sub cmdRvk_Click()
   RevokeAll
   
End Sub


'Grant all in use

Private Sub GrantAll()
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If TypeOf Controls(iList) Is CheckBox Then
         If Controls(iList).Visible Then Controls(iList).Value = vbChecked
      End If
   Next
   
End Sub

'Revoke all in use

Private Sub RevokeAll()
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If TypeOf Controls(iList) Is CheckBox Then
         If Controls(iList).Visible Then Controls(iList).Value = vbUnchecked
      End If
   Next
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      Fillboxes
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   Move AdmnUuser2.Left + 700, AdmnUuser2.Top + 1500
   FormatBoxes
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdmnUperm4 = Nothing
   
End Sub




Private Sub Fillboxes()
   'Note - There are (2) spare groups for Prod, Fina and Invc
   On Error Resume Next
   Get #iFreeDbf, iCurrentRec, Secure
   optEd1.Value = Secure.UserProdG1E
   optVw1.Value = Secure.UserProdG1V
   optFn1.Value = Secure.UserProdG1F
   
   optEd2.Value = Secure.UserProdG2E
   optVw2.Value = Secure.UserProdG2V
   optFn2.Value = Secure.UserProdG2F
   
   optEd3.Value = Secure.UserProdG3E
   optVw3.Value = Secure.UserProdG3V
   optFn3.Value = Secure.UserProdG3F
   
   optEd4.Value = Secure.UserProdG4E
   optVw4.Value = Secure.UserProdG4V
   optFn4.Value = Secure.UserProdG4F
   
   optEd5.Value = Secure.UserProdG5E
   optVw5.Value = Secure.UserProdG5V
   optFn5.Value = Secure.UserProdG5F
   
   optEd6.Value = Secure.UserProdG6E
   optVw6.Value = Secure.UserProdG6V
   optFn6.Value = Secure.UserProdG6F
   
   optEd7.Value = Secure.UserProdG7E
   optVw7.Value = Secure.UserProdG7V
   optFn7.Value = Secure.UserProdG7F
   
   optEd8.Value = Secure.UserProdG8E
   optVw8.Value = Secure.UserProdG8V
   optFn8.Value = Secure.UserProdG8F
   
   'Section
   optGr1.Value = Secure.UserProdG1
   optGr2.Value = Secure.UserProdG2
   optGr3.Value = Secure.UserProdG3
   optGr4.Value = Secure.UserProdG4
   optGr5.Value = Secure.UserProdG5
   optGr6.Value = Secure.UserProdG6
   optGr7.Value = Secure.UserProdG7
   optGr8.Value = Secure.UserProdG8
   
   TestBoxes 1, Secure.UserProdG1
   TestBoxes 2, Secure.UserProdG2
   TestBoxes 3, Secure.UserProdG3
   TestBoxes 4, Secure.UserProdG4
   TestBoxes 5, Secure.UserProdG5
   TestBoxes 6, Secure.UserProdG6
   TestBoxes 7, Secure.UserProdG7
   TestBoxes 8, Secure.UserProdG8
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optEd1_Click()
   Secure.UserProdG1E = optEd1.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optEd1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd2_Click()
   Secure.UserProdG2E = optEd2.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optEd2_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd3_Click()
   Secure.UserProdG3E = optEd3.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optEd3_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd4_Click()
   Secure.UserProdG4E = optEd4.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optEd4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd5_Click()
   Secure.UserProdG5E = optEd5.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optEd5_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd6_Click()
   Secure.UserProdG6E = optEd6.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optEd6_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd7_Click()
   Secure.UserProdG7E = optEd7.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optEd7_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optEd8_Click()
   Secure.UserProdG8E = optEd8.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optEd8_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFn1_Click()
   Secure.UserProdG1F = optFn1.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optFn1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFn2_Click()
   Secure.UserProdG2F = optFn2.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optFn2_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFn3_Click()
   Secure.UserProdG3F = optFn3.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optFn3_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFn4_Click()
   Secure.UserProdG4F = optFn4.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optFn4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFn5_Click()
   Secure.UserProdG5F = optFn5.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optFn5_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFn6_Click()
   Secure.UserProdG6F = optFn6.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optFn6_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFn7_Click()
   Secure.UserProdG7F = optFn7.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optFn7_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFn8_Click()
   Secure.UserProdG8F = optFn8.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optFn8_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr1_Click()
   If optGr1.Value = vbUnchecked Then
      TestBoxes 1, vbUnchecked
      optEd1.Value = vbUnchecked
      optVw1.Value = vbUnchecked
      optFn1.Value = vbUnchecked
      Secure.UserProdG1 = 0
   Else
      TestBoxes 1, vbChecked
      Secure.UserProdG1 = 1
   End If
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optGr1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr2_Click()
   If optGr2.Value = vbUnchecked Then
      TestBoxes 2, vbUnchecked
      optEd2.Value = vbUnchecked
      optVw2.Value = vbUnchecked
      optFn2.Value = vbUnchecked
      Secure.UserProdG2 = 0
   Else
      TestBoxes 2, vbChecked
      Secure.UserProdG2 = 1
   End If
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub optGr2_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr3_Click()
   If optGr3.Value = vbUnchecked Then
      TestBoxes 3, vbUnchecked
      optEd3.Value = vbUnchecked
      optVw3.Value = vbUnchecked
      optFn3.Value = vbUnchecked
      Secure.UserProdG3 = 0
   Else
      TestBoxes 3, vbChecked
      Secure.UserProdG3 = 1
   End If
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub optGr3_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr4_Click()
   If optGr4.Value = vbUnchecked Then
      TestBoxes 4, vbUnchecked
      optEd4.Value = vbUnchecked
      optVw4.Value = vbUnchecked
      optFn4.Value = vbUnchecked
      Secure.UserProdG4 = 0
   Else
      TestBoxes 4, vbChecked
      Secure.UserProdG4 = 1
   End If
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub optGr4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr5_Click()
   If optGr5.Value = vbUnchecked Then
      TestBoxes 5, vbUnchecked
      optEd5.Value = vbUnchecked
      optVw5.Value = vbUnchecked
      optFn5.Value = vbUnchecked
      Secure.UserProdG5 = 0
   Else
      TestBoxes 5, vbChecked
      Secure.UserProdG5 = 1
   End If
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub optGr5_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr6_Click()
   If optGr6.Value = vbUnchecked Then
      TestBoxes 6, vbUnchecked
      optEd6.Value = vbUnchecked
      optVw6.Value = vbUnchecked
      optFn6.Value = vbUnchecked
      Secure.UserProdG6 = 0
   Else
      TestBoxes 6, vbChecked
      Secure.UserProdG6 = 1
   End If
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub optGr6_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr7_Click()
   If optGr7.Value = vbUnchecked Then
      TestBoxes 7, vbUnchecked
      optEd7.Value = vbUnchecked
      optVw7.Value = vbUnchecked
      optFn7.Value = vbUnchecked
      Secure.UserProdG7 = 0
   Else
      TestBoxes 7, vbChecked
      Secure.UserProdG7 = 1
   End If
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub optGr7_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optGr8_Click()
   If optGr8.Value = vbUnchecked Then
      TestBoxes 8, vbUnchecked
      optEd8.Value = vbUnchecked
      optVw8.Value = vbUnchecked
      optFn8.Value = vbUnchecked
      Secure.UserProdG8 = 0
   Else
      TestBoxes 8, vbChecked
      Secure.UserProdG8 = 1
   End If
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub


Private Sub optGr8_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw1_Click()
   Secure.UserProdG1V = optVw1.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optVw1_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw2_Click()
   Secure.UserProdG2V = optVw2.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optVw2_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw3_Click()
   Secure.UserProdG3V = optVw3.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optVw3_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw4_Click()
   Secure.UserProdG4V = optVw4.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optVw4_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw5_Click()
   Secure.UserProdG5V = optVw5.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optVw5_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw6_Click()
   Secure.UserProdG6V = optVw6.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optVw6_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw7_Click()
   Secure.UserProdG7V = optVw7.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optVw7_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optVw8_Click()
   Secure.UserProdG8V = optVw8.Value
   Put #iFreeDbf, iCurrentRec, Secure
   
End Sub

Private Sub optVw8_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub
