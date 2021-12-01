VERSION 5.00
Begin VB.Form PomMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Message for POM Users"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   Icon            =   "PomMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   11310
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H80000018&
      Caption         =   "&Update"
      Height          =   315
      Left            =   10200
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Send Message"
      Top             =   660
      Width           =   875
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   180
      MaxLength       =   512
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "99"
      Top             =   1080
      Width           =   10935
   End
   Begin VB.CommandButton cmdCan 
      BackColor       =   &H80000018&
      Caption         =   "Close"
      Height          =   435
      Left            =   10200
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Close This Dialog (ESC)"
      Top             =   120
      Width           =   875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"PomMessage.frx":08CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   9795
   End
End
Attribute VB_Name = "PomMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub FormatControls()
'   Dim b As Byte
'   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'
'End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdSend_Click()
   sSql = "update Alerts set ALERTMSG = '" & txtMessage & "' where ALERTREF = 1"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected = 0 Then
      sSql = "insert into Alerts(ALERTREF,ALERTMSG)" & vbCrLf _
         & "values( 1, '" & txtMessage & "')"
      clsADOCon.ExecuteSQL sSql
   End If
   
   MsgBox "Message saved"
End Sub


Private Sub Form_Activate()
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   Move 1500, 1500
   'FormatControls
   
   Dim Ado As ADODB.Recordset
   sSql = "select ALERTMSG as Msg from Alerts where ALERTREF = 1"
   If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then
      txtMessage.Text = "" & Ado.Fields(0)
      'Debug.Print "" & rdo.rdoColumns(0)
   End If
   Set Ado = Nothing
End Sub

