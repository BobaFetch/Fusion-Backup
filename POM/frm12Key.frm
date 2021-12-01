VERSION 5.00
Begin VB.Form frm12Key 
   BorderStyle     =   0  'None
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKey 
      Caption         =   "ENTER"
      Enabled         =   0   'False
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
      Index           =   11
      Left            =   0
      TabIndex        =   12
      Top             =   4001
      Width           =   4500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "CLR"
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
      Index           =   12
      Left            =   3000
      TabIndex        =   11
      Top             =   3000
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   10
      Left            =   1501
      TabIndex        =   10
      Top             =   3000
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "0"
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
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   3001
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "1"
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
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   2001
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "2"
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
      Index           =   2
      Left            =   1501
      TabIndex        =   7
      Top             =   2001
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "3"
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
      Index           =   3
      Left            =   3001
      TabIndex        =   6
      Top             =   2001
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "4"
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
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1001
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "5"
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
      Index           =   5
      Left            =   1501
      TabIndex        =   4
      Top             =   1001
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "6"
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
      Index           =   6
      Left            =   3001
      TabIndex        =   3
      Top             =   1001
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "7"
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
      Index           =   7
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "8"
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
      Index           =   8
      Left            =   1501
      TabIndex        =   1
      Top             =   0
      Width           =   1500
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "9"
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
      Index           =   9
      Left            =   3001
      TabIndex        =   0
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "frm12Key"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) ESI 2003
'
'**********************************************************************************
'
' frm12key - Popup keypad input window.
'
' Notes:
'
' Created: (nth) 03/12/03
' Revisions:
'   05/20/03 (nth) Added ability to have a decimal point.
'
'
'**********************************************************************************

Option Explicit

'**********************************************************************************

Private Sub cmdKey_Click(Index As Integer)
   If Index < 11 Then
      If Val(glblActive.Tag) > 0 Then
         If Len(glblActive.Caption) < Val(glblActive.Tag) Then
            glblActive.Caption = glblActive.Caption & cmdKey(Index).Caption
            EnterOn True
         Else
            Beep
         End If
      End If
   ElseIf Index = 12 Then
      ClearInput
   ElseIf Index = 11 Then
      frmMain.ProcessEnter
   End If
End Sub

Public Sub ClearInput()
   glblActive.Caption = ""
   EnterOn False
End Sub

Public Sub EnterOn(pblnON As Boolean)
   cmdKey(11).Enabled = pblnON
End Sub
