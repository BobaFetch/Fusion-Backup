VERSION 5.00
Begin VB.Form LotsLTe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change The User Lot Id"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox lblNumber 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5040
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Update User Lot ID And Exit"
      Top             =   0
      Width           =   875
   End
   Begin VB.TextBox txtLot 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "New User ID (Minimum 10 Chars)"
      Top             =   720
      Width           =   3950
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Lot Number"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Lot Number"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "LotsLTe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit
Dim sOldLot As String

Private Sub cmdCan_Click()
   LotsLTe01a.GetLots 1
   LotsLTe01a.lblNumber = lblNumber
   LotsLTe01a.cmbLot = txtLot
   LotsLTe01a.GetThisLot
   'LotsLTe01a.cmbLot.SetFocus
   Form_Deactivate
   
End Sub


Private Sub Form_Activate()
   sOldLot = txtLot
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   Move 2000, 2000
   lblNumber.BackColor = BackColor
   lblNumber = LotsLTe01a.lblNumber
   LotsLTe01a.cmbPrt.Enabled = False
   
End Sub

Private Sub txtLot_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub txtLot_KeyPress(KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtlot_LostFocus()
   Dim bByte As Byte
   txtLot = CheckLen(txtLot, 40)
   If Trim(txtLot) <> sOldLot Then
      If Len(Trim(txtLot)) < 5 Then
         Beep
         txtLot = sOldLot
         MsgBox "New User Lots Require At Least (5 chars).", _
            vbInformation
      Else
         bByte = GetUserLotID(Trim(txtLot))
         If bByte = 0 Then
            sSql = "UPDATE LohdTable SET LOTUSERLOTID='" & txtLot & "' " _
                   & "WHERE LOTNUMBER='" & lblNumber & "'"
            clsADOCon.ExecuteSQL sSql
         Else
            txtLot = sOldLot
         End If
      End If
   End If
   sOldLot = txtLot
   LotsLTe01a.cmbLot = txtLot
   
End Sub
