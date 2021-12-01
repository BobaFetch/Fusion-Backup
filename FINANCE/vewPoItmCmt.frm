VERSION 5.00
Begin VB.Form vewPoItmCmt 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "vewPoItmCmt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCmt 
      Height          =   3375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Tag             =   "9"
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "vewPoItmCmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bOnLoad As Byte
Const FRMHEIGHT As Integer = 3375

Private Sub Form_Activate()
   If bOnLoad = 1 Then
      Dim i As Integer
      For i = 0 To FRMHEIGHT Step 10
         Me.Height = i
         DoEvents
      Next
   End If
   bOnLoad = 0
End Sub

Private Sub Form_Click()
   Unload Me
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_Load()
   'SetFormSize Me
   Me.Height = 0
   bOnLoad = 1
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   WindowState = 1
   Set VewPsItm = Nothing
End Sub

Private Sub txtCmt_KeyPress(KeyAscii As Integer)
   KeyCase KeyAscii
   If KeyAscii = 27 Then Unload Me
   
End Sub
