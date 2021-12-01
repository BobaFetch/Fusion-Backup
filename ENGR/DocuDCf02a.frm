VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Document"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDoc 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "9"
      ToolTipText     =   "Contains Parts With A Document List"
      Top             =   720
      Width           =   3285
   End
   Begin VB.ComboBox cboSheet 
      Height          =   315
      Left            =   4560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "9"
      ToolTipText     =   "Sheet (If Marked In Class)"
      Top             =   1140
      Width           =   735
   End
   Begin VB.ComboBox cboRev 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "9"
      ToolTipText     =   "Document Revision"
      Top             =   1140
      Width           =   1215
   End
   Begin VB.ComboBox cboClass 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "DocuDCf02a.frx":0000
      Left            =   2040
      List            =   "DocuDCf02a.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "9"
      ToolTipText     =   "Select Class From List"
      Top             =   300
      Width           =   2000
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCf02a.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   5580
      TabIndex        =   5
      ToolTipText     =   "Delete This Revision"
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txtDsc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   1620
      Width           =   4200
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5580
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2295
      FormDesignWidth =   6750
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Number"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   1725
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sheet"
      Height          =   255
      Index           =   3
      Left            =   3780
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Class"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   300
      Width           =   1725
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   1620
      Width           =   1215
   End
End
Attribute VB_Name = "DocuDCf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'5/14/04 New
Option Explicit
Dim bOnLoad As Byte
Dim sOldDoc As String
Private doc As ClassDoc

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Private Sub cmbDoc_Click()
'   FillRevisions
'
'End Sub
'
'
'Private Sub cmbRev_Click()
'   FillSheets
'
'End Sub
'
'
Private Sub cboClass_Click()
   doc.FillDocuments cboClass, cboDoc
End Sub

Private Sub cboClass_LostFocus()
   doc.FillDocuments cboClass, cboDoc
End Sub

Private Sub cboDoc_Click()
   doc.FillRevisions cboClass, cboDoc, cboRev
   txtDsc = doc.GetDocDesc(cboClass, cboDoc)
End Sub

Private Sub cboDoc_LostFocus()
   doc.FillRevisions cboClass, cboDoc, cboRev
   txtDsc = doc.GetDocDesc(cboClass, cboDoc)
End Sub

Private Sub cboRev_Click()
   doc.FillSheets cboClass, cboDoc, cboRev, cboSheet
   txtDsc = doc.GetDocDesc(cboClass, cboDoc)
End Sub

Private Sub cboRev_LostFocus()
   doc.FillSheets cboClass, cboDoc, cboRev, cboSheet
   txtDsc = doc.GetDocDesc(cboClass, cboDoc)
End Sub

Private Sub cboSheet_Click()
   txtDsc = doc.GetDocDesc(cboClass, cboDoc)
End Sub

Private Sub cboSheet_LostFocus()
   txtDsc = doc.GetDocDesc(cboClass, cboDoc)
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdDel_Click()
   If txtDsc.ForeColor = ES_RED Then
      MsgBox "You Must Select A Valid Document.", _
         vbInformation, Caption
   Else
      DeleteDocument
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3351
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      'FillCombo
      Set doc = New ClassDoc
      doc.FillClasses cboClass, False
      bOnLoad = 0
      MouseCursor 0
   End If
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set DocuDCf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = Es_FormBackColor
   sOldDoc = ""
   
End Sub

'Private Sub FillCombo()
'   On Error GoTo DiaErr1
'   cmbDoc.Clear
'   sSql = "Qry_FillDocuments"
'   LoadComboBox cmbDoc
'   If cmbDoc.ListCount > 0 Then
'      cmbDoc = cmbDoc.List(0)
'      ' FillRevisions
'   Else
'      cmdDel.Enabled = False
'      MsgBox "No Documents Are Recorded.", _
'         vbInformation, Caption
'   End If
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'Private Sub FillRevisions()
'   Dim RdoRev As ADODB.Recordset
'   On Error GoTo DiaErr1
'   If sOldDoc <> Trim(cmbDoc) Then
'      cmbRev.Clear
'      sSql = "SELECT DISTINCT DOREV FROM DdocTable WHERE " _
'             & "DOREF='" & Compress(cmbDoc) & "' " _
'             & "ORDER BY DOREV"
'      bSqlRows = clsADOCon.GetDataSet(sSql,RdoRev, ES_FORWARD)
'      If bSqlRows Then
'         With RdoRev
'            Do Until .EOF
'               cmbRev.AddItem !DOREV
'               .MoveNext
'            Loop
'            ClearResultSet RdoRev
'         End With
'      End If
'      If cmbRev.ListCount > 0 Then cmbRev = cmbRev.List(0)
'      GetDocument
'   End If
'   sOldDoc = Trim(cmbDoc)
'   Set RdoRev = Nothing
'   FillSheets
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillrevisions"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'Private Sub FillSheets()
'   Dim RdoSht As ADODB.Recordset
'   On Error GoTo DiaErr1
'   cmbSht.Clear
'   sSql = "SELECT DISTINCT DOSHEET FROM DdocTable WHERE " _
'          & "DOREF='" & Compress(cmbDoc) & "' AND " _
'          & "DOREV='" & Trim(cmbRev) & "' ORDER BY DOSHEET"
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoSht, ES_FORWARD)
'   If bSqlRows Then
'      With RdoSht
'         Do Until .EOF
'            cmbSht.AddItem !DOSHEET
'            .MoveNext
'         Loop
'         ClearResultSet RdoSht
'      End With
'   End If
'   If cmbSht.ListCount > 0 Then cmbSht = cmbSht.List(0)
'   GetDocument
'   Set RdoSht = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillsheets"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'Private Sub GetDocument()
'   Dim RdoDoc As ADODB.Recordset
'   On Error GoTo DiaErr1
'   sSql = "SELECT DODESCR FROM DdocTable WHERE (DOREF='" _
'          & Compress(cmbDoc) & "' AND DOREV='" _
'          & Trim(cmbRev) & "' AND DOSHEET='" & Trim(cmbSht) & "')"
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoDoc, ES_FORWARD)
'   If bSqlRows Then
'      With RdoDoc
'         If Not IsNull(.rdoColumns(0)) Then _
'                       txtDsc = Trim(.rdoColumns(0)) _
'                       Else txtDsc = "*** Document Wasn't Found ***"
'         ClearResultSet RdoDoc
'      End With
'   Else
'      txtDsc = "*** Document Wasn't Found ***"
'   End If
'   Set RdoDoc = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getdocument"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
'Private Sub txtDsc_Change()
'   If Left(txtDsc, 8) = "*** Docu" Then
'      txtDsc.ForeColor = ES_RED
'      cmdDel.Enabled = False
'   Else
'      txtDsc.ForeColor = vbBlack
'      cmdDel.Enabled = True
'   End If
'
'End Sub
'
'

Private Sub DeleteDocument()
   Dim RdoDel As ADODB.Recordset
   Dim bResponse As Byte
   Dim iCount As Integer
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sMsg = "This Procedure Will Permanently Remove The " & vbCrLf _
          & "Document From The List And Cannot Be " & vbCrLf _
          & "Reversed. Continue To Delete This Document?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      'Test It
      'Lists
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      bResponse = 0
      sSql = "SELECT DLSREF,DLSDOCREF,DLSDOCREV,DLSDOCSHEET FROM DlstTable" & vbCrLf _
         & "WHERE DLSDOCCLASS = '" & Compress(cboClass) & "' AND DLSDOCREF='" & Compress(cboDoc) _
             & "' AND DLSDOCREV='" & Trim(cboRev) & "' AND DLSDOCSHEET='" _
             & Trim(cboSheet) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoDel, ES_FORWARD)
      If bSqlRows Then
         With RdoDel
            Do Until .EOF
               iCount = iCount + 1
               .MoveNext
            Loop
            ClearResultSet RdoDel
         End With
         If iCount > 0 Then
            clsADOCon.RollbackTrans
            MsgBox "This Document Is Used On " & iCount & " " _
               & "Document List(s) And Cannot Be Deleted.", _
               vbInformation, Caption
            bResponse = 1
         Else
            iCount = 0
            'MO lists
            sSql = "SELECT RUNDLSNUM,RUNDLSDOCREF,RUNDLSDOCREV,RUNDLSDOCREFSHEET " _
                   & "FROM RndlTable WHERE (RUNDLSDOCREF='" & Compress(cboDoc) _
                   & "' AND RUNDLSDOCREV='" & Trim(cboRev) & "' AND RUNDLSDOCREFSHEET='" _
                   & Trim(cboSheet) & "')"
            bSqlRows = clsADOCon.GetDataSet(sSql, RdoDel, ES_FORWARD)
            If bSqlRows Then
               With RdoDel
                  Do Until .EOF
                     iCount = iCount + 1
                     .MoveNext
                  Loop
                  ClearResultSet RdoDel
               End With
            End If
            If iCount > 0 Then
               clsADOCon.RollbackTrans
               MsgBox "This Document Is Used On " & iCount & " " _
                  & "MO Document List(s) And Cannot Be Deleted.", _
                  vbInformation, Caption
               bResponse = 1
            End If
         End If
         
      End If
   Else
      bResponse = 1
      CancelTrans
   End If
   Set RdoDel = Nothing
   If bResponse = 0 Then
      'On Error Resume Next
      'clsADOCon.BeginTrans
      sSql = "DELETE FROM DdocTable" & vbCrLf _
         & "WHERE DOCLASS = '" & Compress(cboClass) & "' AND DOREF='" _
         & Compress(cboDoc) & "' AND DOREV='" _
         & Trim(cboRev) & "' AND DOSHEET='" _
         & Trim(cboSheet) & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.ADOErrNum > 0 Then
         clsADOCon.RollbackTrans
         MsgBox "Could Not Delete The Document.", _
            vbExclamation, Caption
      Else
         clsADOCon.CommitTrans
         SysMsg "The Document Was Deleted.", True
         'FillCombo
         doc.FillDocuments cboClass, cboDoc
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "deletedocum"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
