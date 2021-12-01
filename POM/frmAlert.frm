VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "pstrCaption"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2160
      TabIndex        =   0
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Image imgQuestion 
      Height          =   480
      Left            =   120
      Picture         =   "frmAlert.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgInfo 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "frmAlert.frx":0742
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgHalt 
      Height          =   480
      Left            =   120
      Picture         =   "frmAlert.frx":0A4C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Caption         =   "pstrMsg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) ESI 2003
'
'**********************************************************************************
'
' frmAlert
'
' Notes: Send system (MsgBox) messages in a larger format
'
' Created: 03/25/03
' Revisions:
'
'
'**********************************************************************************

Option Explicit


'**********************************************************************************

Private Sub cmdOK_Click()
   gintResponse = vbOK
   Unload Me
End Sub

Private Sub cmdNo_Click()
   gintResponse = vbNo
   Unload Me
End Sub

Private Sub cmdYes_Click()
   gintResponse = vbYes
   Unload Me
End Sub

Private Sub Form_Activate()
   If cmdOK.Visible = False And cmdYes.Visible = False Then
      Refresh
      Sleep 2000
      Unload Me
   End If
End Sub

