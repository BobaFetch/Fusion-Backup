VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy a Document"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3301
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Document"
      Height          =   495
      Left            =   4620
      TabIndex        =   11
      ToolTipText     =   "Edit the new document"
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Document"
      Height          =   495
      Left            =   2940
      TabIndex        =   10
      ToolTipText     =   """Create the new document"""
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopyInfo 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4260
      TabIndex        =   4
      ToolTipText     =   "Copy parameters from above to below"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Copy To"
      Height          =   1995
      Left            =   240
      TabIndex        =   21
      Top             =   3180
      Width           =   8595
      Begin VB.TextBox txtToSheet 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtToRevision 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtToDocDesc 
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtToDocument 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cboToClass 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "DocuDCf03a.frx":0000
         Left            =   1440
         List            =   "DocuDCf03a.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "9"
         ToolTipText     =   "Select Class From List"
         Top             =   360
         Width           =   2000
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Document "
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblToClassDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4920
         TabIndex        =   25
         Top             =   360
         Width           =   3540
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Document Class"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Revision"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copy From:"
      Height          =   1995
      Left            =   180
      TabIndex        =   14
      Top             =   600
      Width           =   8595
      Begin VB.ComboBox cboFromRev 
         Height          =   315
         ItemData        =   "DocuDCf03a.frx":0004
         Left            =   1440
         List            =   "DocuDCf03a.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "9"
         ToolTipText     =   "Document Revision"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cboFromSheet 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "9"
         ToolTipText     =   "Sheet (If Marked In Class)"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cboFromClass 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "DocuDCf03a.frx":0008
         Left            =   1440
         List            =   "DocuDCf03a.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "9"
         ToolTipText     =   "Select Class From List"
         Top             =   360
         Width           =   2000
      End
      Begin VB.ComboBox cboFromDoc 
         Height          =   315
         ItemData        =   "DocuDCf03a.frx":000C
         Left            =   1440
         List            =   "DocuDCf03a.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "9"
         ToolTipText     =   "Enter/Revise Document"
         Top             =   720
         Width           =   3345
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Revision"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblFromDocDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4920
         TabIndex        =   18
         Top             =   720
         Width           =   3540
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Document Class"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblFromClassDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4920
         TabIndex        =   16
         Top             =   360
         Width           =   3540
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Document "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCf03a.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7860
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
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
      FormDesignHeight=   5925
      FormDesignWidth =   8970
   End
End
Attribute VB_Name = "DocuDCf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'''*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'''*** and is protected under US and International copyright    ***
'''*** laws and treaties.                                       ***
'9/12/08 cmbfromclass changed from 8 to blank
Option Explicit
Dim bOnLoad As Byte
Dim bGoodDoc As Byte
Private sClass As String
Private sDoc As String
Private sRev As String
Private sSheet As String
Private updatingCombos As Boolean     'true when in UpdateCombos to avoid re-entry

Private doc As New ClassDoc

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboFromClass_Click()
   UpdateFromCombos
End Sub

Private Sub cboFromClass_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFromClass_LostFocus()
   UpdateFromCombos
End Sub

Private Sub cboFromDoc_Click()
   UpdateFromCombos
End Sub

Private Sub cboFromDoc_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFromDoc_LostFocus()
   UpdateFromCombos
End Sub

Private Sub cboFromRev_Click()
   UpdateFromCombos
End Sub

Private Sub cboFromRev_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFromRev_LostFocus()
   UpdateFromCombos
End Sub

Private Sub cboFromSheet_Change()
   UpdateFromCombos
End Sub

Private Sub cboFromSheet_Click()
   UpdateFromCombos
End Sub

Private Sub cboToClass_Click()
   'lblToClassDesc = GetClassDesc(cboToClass)
   'Dim doc As New ClassDoc
   lblToClassDesc = doc.GetClassDesc(cboToClass)
End Sub

Private Sub cboToClass_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboToClass_LostFocus()
   'Dim doc As New ClassDoc
   lblToClassDesc = doc.GetClassDesc(cboToClass)
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCopyInfo_Click()
   If cboToClass <> cboFromClass Then
      Dim i As Integer
      For i = 0 To cboToClass.ListCount - 1
         If cboToClass.List(i) = cboFromClass Then
            cboToClass.ListIndex = i
            Exit For
         End If
      Next
   End If
   txtToDocument = cboFromDoc
   txtToDocDesc = lblFromDocDesc
   txtToRevision = cboFromRev
   txtToSheet = cboFromSheet
End Sub

Private Sub cmdCreate_Click()
   If Trim(cboToClass) = "" Then
      MsgBox "Destination document class must be specified"
      Exit Sub
   End If
   
    If Trim(Me.txtToDocument) = "" Then
      MsgBox "Destination document must be specified"
      Exit Sub
   End If
   
   If DocExists Then
      MsgBox "That document already exists"
      Exit Sub
   End If
      
   MouseCursor ccHourglass
   cmdCreate.Enabled = False
   CreateDoc
   cmdCreate.Enabled = True
   MouseCursor ccDefault
End Sub

Private Sub CreateDoc()

    Dim sql1 As String, sql2 As String
    
   On Error GoTo whoops
   Dim strUser As String
   strUser = Secure.UserInitials
   
   strUser = StripControlChars(strUser)
   
   sql1 = "insert into DdocTable(DOCLASS,DOREF,DOREV,DOSHEET,DONUM,DODESCR,DOUSER," & vbCrLf _
      & "DOEXTDESC,DOMIC,DOSIZE,DOECO,DOLOC,DORECTYPE,DOADCN,DOTYPE,DOFILENAME,DONOTES,DOCUST)" & vbCrLf _
      & "select '" & Compress(cboToClass) & "','" & Compress(txtToDocument) & "','" & txtToRevision & "'," _
      & "'" & Me.txtToSheet & "','" & txtToDocument & "','" & Replace(txtToDocDesc, "'", "''") & "','" & strUser & "',"
    
    sql2 = "DOEXTDESC,DOMIC,DOSIZE,DOECO,DOLOC,DORECTYPE,DOADCN,DOTYPE,DOFILENAME,DONOTES,DOCUST" & vbCrLf _
      & "FROM DdocTable" & vbCrLf _
      & "WHERE DOCLASS = '" & Compress(sClass) & "' AND DOREF = '" & Compress(sDoc) & "'" & vbCrLf _
      & "AND DOREV = '" & sRev & "' AND DOSHEET = '" & sSheet & "'"
      
   Dim sql As String, doc As String
   
   sSql = sql1 & sql2
   sql = sSql
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   doc = "Document class '" & Compress(cboToClass) & "' number '" & txtToDocument _
      & "' rev '" & txtToRevision & "' sheet '" & txtToSheet & "'"
   If clsADOCon.RowsAffected > 0 Then
      MsgBox doc & " created."
   Else
      MsgBox doc & " creation failed.  " & vbCrLf & sql
   End If
   Exit Sub
   
whoops:
   sProcName = "CreateDoc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Function StripControlChars(source As String, Optional KeepCRLF As Boolean = _
    True) As String
    Dim Index As Long
    Dim bytes() As Byte
    
    ' the fastest way to process this string
    ' is copy it into an array of Bytes
    bytes() = source
    For Index = 0 To UBound(bytes) Step 2
        ' if this is a control character
        If bytes(Index) < 32 And bytes(Index + 1) = 0 Then
            If Not KeepCRLF Or (bytes(Index) <> 13 And bytes(Index) <> 10) Then
                ' the user asked to trim CRLF or this
                ' character isn't a CR or a LF, so clear it
                bytes(Index) = 0
            End If
        End If
    Next
    
    ' return this string, after filtering out all null chars
    StripControlChars = Replace(bytes(), vbNullChar, "")
            
End Function

Private Sub cmdEdit_Click()
   If DocExists Then
'      Dim document As New DocuDCe01a
'      document.PassedClass = cboToClass
'      document.PassedDoc = txtToDocument
'      document.PassedRev = txtToRevision
'      document.PassedSheet = txtToSheet
'      document.Show
      DocuDCe01a.PassedClass = cboToClass
      DocuDCe01a.PassedDoc = txtToDocument
      DocuDCe01a.PassedRev = txtToRevision
      DocuDCe01a.PassedSheet = txtToSheet
      DocuDCe01a.Show
      
      Unload Me
   Else
      MsgBox "You must create the document first"
      Exit Sub
   End If
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      sClass = "!^&%"
      Set doc = New ClassDoc
      doc.FillClasses cboFromClass, False
      lblFromClassDesc = doc.GetClassDesc(cboFromClass)
      doc.FillDocuments cboFromClass, cboFromDoc
      
      doc.FillClasses cboToClass, True
      
      UpdateFromCombos
      bOnLoad = 0
      MouseCursor 0
   End If

End Sub

Private Function GetDocDesc() As String
   Dim rdo As ADODB.Recordset
   sSql = "SELECT DODESCR" & vbCrLf _
      & "FROM DdocTable" & vbCrLf _
      & "WHERE DOCLASS = '" & sClass & "' AND DOREF = '" & sDoc & "'" & vbCrLf _
      & "AND DOREV = '" & sRev & "' AND DOSHEET = '" & sSheet & "'"
   If clsADOCon.GetDataSet(sSql, rdo, ES_KEYSET) Then
      With rdo
         GetDocDesc = Trim("" & !DODESCR)
         If GetDocDesc = "" Then
            GetDocDesc = "<NONE>"
         End If
         ClearResultSet rdo
      End With
   End If
   Set rdo = Nothing

End Function

Private Sub Form_Load()
   FormLoad Me
   bOnLoad = 1
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub UpdateFromCombos()

   If updatingCombos Then
      Exit Sub
   Else
      updatingCombos = True
   End If

   Dim rippleDown As Boolean
   If Compress(cboFromClass) <> sClass Then
      sClass = Compress(cboFromClass)
      rippleDown = True
      'Dim doc As New ClassDoc
      lblFromClassDesc = doc.GetClassDesc(cboFromClass)
      doc.FillDocuments cboFromClass, cboFromDoc
   End If
   
   If Compress(cboFromDoc) <> sDoc Or rippleDown Then
      sDoc = Compress(cboFromDoc)
      rippleDown = True
      doc.FillRevisions cboFromClass, cboFromDoc, cboFromRev
   End If
      
   If Compress(cboFromRev) <> sRev Or rippleDown Then
      sRev = Compress(cboFromRev)
      rippleDown = True
      doc.FillSheets cboFromClass, cboFromDoc, cboFromRev, cboFromSheet
   End If
      
   sSheet = cboFromSheet
   lblFromDocDesc = GetDocDesc
   updatingCombos = False
End Sub

Private Sub txtToDocument_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtToRevision_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function DocExists() As Boolean
   Dim rdo As ADODB.Recordset
   sSql = "SELECT DODESCR" & vbCrLf _
      & "FROM DdocTable" & vbCrLf _
      & "WHERE DOCLASS = '" & Compress(cboToClass) & "' AND DOREF = '" & Compress(txtToDocument) & "'" & vbCrLf _
      & "AND DOREV = '" & txtToRevision & "' AND DOSHEET = '" & txtToSheet & "'"
   If clsADOCon.GetDataSet(sSql, rdo, ES_KEYSET) Then
      DocExists = True
   End If
   Set rdo = Nothing
End Function
