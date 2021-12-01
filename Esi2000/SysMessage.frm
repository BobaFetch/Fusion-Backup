VERSION 5.00
Begin VB.Form SysMessage 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Broadcast Message"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "SysMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMessage 
      Height          =   1365
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton cmdCan 
      BackColor       =   &H80000018&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Close This Dialog (ESC)"
      Top             =   120
      Width           =   875
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Broadcast A Message To All Logons"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "SysMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdSend_Click()
   If Trim(txtFrom) = "" Then
      MsgBox "Requires From.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Trim(txtSubject) = "" Then
      MsgBox "Requires Subject.", _
         vbInformation, Caption
      Exit Sub
   End If
   If Trim(txtMessage) = "" Then
      MsgBox "Requires A Message.", _
         vbInformation, Caption
      Exit Sub
   End If
   SendThisMessage
   
End Sub


Private Sub Form_Activate()
   MouseCursor 0
   
End Sub

Private Sub Form_Initialize()
   BackColor = ES_ViewBackColor
   z1(3).ForeColor = ES_BLUE
   
End Sub

Private Sub Form_Load()
   '
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SysMessage = Nothing
   
End Sub


Private Sub txtFrom_LostFocus()
   txtFrom = CheckLen(txtFrom, 20)
   txtFrom = StrCase(txtFrom)
   
End Sub


Private Sub txtMessage_LostFocus()
   txtMessage = CheckLen(txtMessage, 255)
   
End Sub


Private Sub txtSubject_LostFocus()
   txtSubject = CheckLen(txtSubject, 30)
   txtSubject = StrCase(txtSubject)
   
End Sub



Private Sub SendThisMessage()
   Dim RdoMsg As ADODB.Recordset
   Dim lMsgID As Long

   On Error GoTo DiaErr1
   clsADOCon.ExecuteSQL "use msdb"
   txtTime = Format(Now, "mm/dd/yy hh:mm AMPM")
   sSql = "SELECT MAX(Message_ID) as ID FROM SystemMessages"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMsg, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoMsg!ID) Then
         lMsgID = RdoMsg!ID + 1
      Else
         lMsgID = 1
      End If
   Else
      lMsgID = 1
   End If
   sSql = "INSERT INTO SystemMessages(Message_ID,Message_TIME," _
          & "Message_FROM,Message_HEADER,Message_TEXT) VALUES (" _
          & lMsgID & ",'" & txtTime & "','" & txtFrom & "','" _
          & txtSubject & "','" & txtMessage & "')"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   z1(3).Caption = "Message Sent..."
   z1(3).Refresh
   Sleep 3000
   clsADOCon.ExecuteSQL "use " & sDataBase
   Unload Me
   Exit Sub
DiaErr1:
   On Error Resume Next
   clsADOCon.ExecuteSQL "use " & sDataBase
   MsgBox "Persons Sending Messages Must Have Read/Write " & vbCr _
      & "Permissions to msdb.  Persons Receiving Messages " & vbCr _
      & "Must Have Read Permissions To msdb At A Minimum.", _
      vbExclamation, Caption
   Unload Me
   
End Sub

Private Sub Document()
   sSql = "CREATE TABLE SystemMessages (" _
          & "Message_ID INT NULL DEFAULT(0)," _
          & "Message_TIME CHAR(20) NULL DEFAULT('')," _
          & "Message_FROM CHAR(20) NULL DEFAULT('')," _
          & "Message_HEADER CHAR(30) NULL DEFAULT('')," _
          & "Message_TEXT VARCHAR(255) NULL DEFAULT(''))"
   
End Sub
