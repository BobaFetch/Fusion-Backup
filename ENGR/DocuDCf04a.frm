VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete a Document Class"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3301
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboClass 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "DocuDCf04a.frx":0000
      Left            =   480
      List            =   "DocuDCf04a.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Class From List"
      Top             =   600
      Width           =   2000
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   480
      Left            =   2340
      TabIndex        =   1
      Top             =   1080
      Width           =   900
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCf04a.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   480
      Left            =   3600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   900
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   5700
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1830
      FormDesignWidth =   6870
   End
   Begin VB.Label lblClassDesc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   3540
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select an Empty Document Class to Delete:"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   5235
   End
End
Attribute VB_Name = "DocuDCf04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'''*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'''*** and is protected under US and International copyright    ***
'''*** laws and treaties.                                       ***

Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub FillClasses(cbo As ComboBox)

   sSql = "SELECT DCLNAME FROM DclsTable" & vbCrLf _
      & "WHERE DCLREF NOT IN (SELECT DOCLASS FROM DdocTable)" & vbCrLf _
      & "ORDER BY DCLNAME"

   LoadComboBox cbo, -1
   If cbo.ListCount = 0 Then
      MsgBox "There are no empty document classes to delete", vbExclamation, Caption
      Unload Me
   End If
End Sub

Private Sub cboClass_Click()
   Dim doc As New ClassDoc
   lblClassDesc = doc.GetClassDesc(cboClass)
End Sub

'Private Sub cboClass_KeyPress(KeyAscii As Integer)
'   KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub cboClass_LostFocus()
   Dim doc As New ClassDoc
   lblClassDesc = doc.GetClassDesc(cboClass)
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   If cboClass = "" Then
      MsgBox "No class selected to delete"
      Exit Sub
   End If
   
   Select Case MsgBox("Delete document class '" & cboClass & "'?", vbYesNo)
   Case vbYes
      MouseCursor ccHourglass
      cmdDelete.Enabled = False
      DeleteClass cboClass.Text
      FillClasses cboClass
      cmdDelete.Enabled = True
      MouseCursor ccDefault
   End Select
End Sub

Private Sub DeleteClass(ClassName As String)
   Dim doc As New ClassDoc
   
   If doc.DeleteClass(ClassName) Then
      MsgBox "Document class '" & ClassName & "' deleted."
   End If
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillClasses cboClass
      bOnLoad = 0
      MouseCursor 0
   End If
End Sub

Private Sub Form_Load()
   FormLoad Me
   bOnLoad = 1
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

