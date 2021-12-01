VERSION 5.00
Begin VB.Form HelpLink
   BackColor = &H8000000C&
   BorderStyle = 1 'Fixed Single
   Caption = "Object Linking"
   ClientHeight = 4452
   ClientLeft = 48
   ClientTop = 336
   ClientWidth = 4260
   Icon = "HelpLink.frx":0000
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4452
   ScaleWidth = 4260
   Begin VB.Label lblForm
      Height = 255
      Left = 240
      TabIndex = 1
      Top = 4320
      Visible = 0 'False
      Width = 375
   End
   Begin VB.Label Label1
      Appearance = 0 'Flat
      BackColor = &H80000005&
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      ForeColor = &H80000008&
      Height = 4215
      Left = 120
      TabIndex = 0
      Top = 120
      Width = 4065
   End
End
Attribute VB_Name = "HelpLink"
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

Private Sub Form_Unload(Cancel As Integer)
   If Val(lblForm) = 0 Then
      DocuDCe04a.optHelp.Value = vbUnchecked
   Else
      DocuDCe04a.optHelp.Value = vbUnchecked
   End If
   
End Sub
