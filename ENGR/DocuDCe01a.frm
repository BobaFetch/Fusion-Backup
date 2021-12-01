VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documents"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3301
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFromDocLst 
      Height          =   255
      Left            =   4080
      TabIndex        =   48
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdDockLok 
      Height          =   495
      Left            =   6120
      Picture         =   "DocuDCe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Push Info to DocumentLok"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdNew 
      Cancel          =   -1  'True
      Caption         =   "New"
      Height          =   435
      Left            =   5940
      TabIndex        =   4
      Top             =   1140
      Width           =   875
   End
   Begin VB.Frame fraEdit 
      Height          =   5115
      Left            =   120
      TabIndex        =   29
      Top             =   1740
      Width           =   6855
      Begin VB.TextBox txtDsc 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Tag             =   "2"
         Top             =   180
         Width           =   5160
      End
      Begin VB.TextBox txtExt 
         Height          =   735
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "9"
         ToolTipText     =   "Comments (2048 Chars Max)"
         Top             =   540
         Width           =   3495
      End
      Begin VB.TextBox txtMic 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Tag             =   "3"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox txtSze 
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Tag             =   "3"
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox txtLoc 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Tag             =   "3"
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   4080
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1740
         Width           =   552
      End
      Begin VB.TextBox txtAdc 
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Tag             =   "3"
         Top             =   2100
         Width           =   2535
      End
      Begin VB.TextBox txtTyp 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Tag             =   "3"
         Top             =   2460
         Width           =   2025
      End
      Begin VB.ComboBox cmbCst 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   19
         Tag             =   "8"
         ToolTipText     =   "Select Customer From List"
         Top             =   3900
         Width           =   1555
      End
      Begin VB.TextBox txtNte 
         Height          =   615
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Tag             =   "9"
         ToolTipText     =   "Comments (2048 Chars Max)"
         Top             =   4380
         Width           =   5175
      End
      Begin VB.TextBox txtEco 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Tag             =   "1"
         Top             =   2100
         Width           =   435
      End
      Begin VB.TextBox txtFle 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Tag             =   "2"
         Top             =   2820
         Width           =   2295
      End
      Begin VB.ComboBox txtRdte 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Tag             =   "4"
         Top             =   3180
         Width           =   1250
      End
      Begin VB.ComboBox txtAdte 
         Height          =   315
         Left            =   4080
         TabIndex        =   16
         Tag             =   "4"
         Top             =   3180
         Width           =   1250
      End
      Begin VB.ComboBox txtEdte 
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Tag             =   "4"
         Top             =   3540
         Width           =   1250
      End
      Begin VB.ComboBox txtOdte 
         Height          =   315
         Left            =   4080
         TabIndex        =   18
         Tag             =   "4"
         Top             =   3540
         Width           =   1250
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext Description"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Microfilm Loc"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   43
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   42
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   255
         Index           =   9
         Left            =   2760
         TabIndex        =   41
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "ADCN's"
         Height          =   255
         Index           =   10
         Left            =   2760
         TabIndex        =   40
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   39
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Received"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   38
         Top             =   3180
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "ADCN Date"
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   37
         Top             =   3180
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   36
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Obsolete Date"
         Height          =   255
         Index           =   15
         Left            =   2880
         TabIndex        =   35
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   34
         Top             =   3900
         Width           =   1215
      End
      Begin VB.Label txtNme 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3120
         TabIndex        =   33
         Top             =   3900
         Width           =   3495
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   32
         Top             =   4380
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "ECO"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   31
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   30
         Top             =   2820
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCe01a.frx":05DA
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "DocuDCe01a.frx":0D88
      Height          =   315
      Left            =   5520
      Picture         =   "DocuDCe01a.frx":1262
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Show/Update Document Lists"
      Top             =   0
      Width           =   350
   End
   Begin VB.ComboBox cboSheet 
      Height          =   315
      Left            =   3700
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Sheet (If Marked In Class)"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cboRev 
      Height          =   315
      ItemData        =   "DocuDCe01a.frx":173C
      Left            =   1560
      List            =   "DocuDCe01a.frx":173E
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Document Revision"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cboDoc 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise Document"
      Top             =   960
      Width           =   3345
   End
   Begin VB.ComboBox cboClass 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "9"
      ToolTipText     =   "Select Class From List"
      Top             =   600
      Width           =   2000
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   21
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2760
      Top             =   60
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6945
      FormDesignWidth =   7110
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sheet"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   26
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   25
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document "
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   24
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3600
      TabIndex        =   23
      Top             =   600
      Width           =   3240
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Class"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "DocuDCe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/3/05 Added DocuDCe01b to update Document Lists
Option Explicit
Dim RdoDoc As ADODB.Recordset

Dim bCancel As Byte
Dim bOnLoad As Byte
'Dim bGoodDoc As Boolean
Dim bDataChanged As Byte

Dim sOldDoc As String
Dim sOldCls As String
Private doc As ClassDoc
Private sClass As String
Private sDoc As String
Private sRev As String
Private sSheet As String
Private updatingCombos As Boolean

'parameters passed from docudcf03a
Public PassedClass As String
Public PassedDoc As String
Public PassedRev As String
Public PassedSheet As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cboClass_Click()
   UpdateCombos
End Sub


Private Sub cboClass_LostFocus()
   If bCancel = 1 Then Exit Sub
   UpdateCombos
   ExitCombos
End Sub


Private Sub cmbCst_Click()
   If cmbCst <> "NONE" Then
      FindCustomer Me, cmbCst, False
   Else
      txtNme = "Not Customer Specific"
   End If
   
End Sub

Private Sub cmbCst_LostFocus()
   Dim sCust As String
   sCust = Compress(cmbCst)
   If Len(Trim(cmbCst)) = 0 Then
      cmbCst = "NONE"
      txtNme = "Not Customer Specific"
   End If
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DOCUST = "" & sCust
      If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub

Private Sub cboDoc_Click()
   'GetRevisions
   UpdateCombos
End Sub

Private Sub cboDoc_LostFocus()
   If bCancel = 1 Then Exit Sub
   UpdateCombos
'   cboDoc = CheckLen(cboDoc, 30)
'   bGoodDoc = GetThisDocument()
'   'If bGoodDoc Then
'      GetRevisions
'   'End If
   ExitCombos
End Sub

Private Sub cboRev_Click()
'   If bCancel = 1 Then Exit Sub
'   If cboSheet.Visible Then GetSheets
'   bGoodDoc = GetThisDocument()
   UpdateCombos
End Sub

Private Sub cboRev_LostFocus()
'   cboRev = CheckLen(cboRev, 6)
'   If cboSheet.Visible = False Then
'      cboSheet = ""
'      bGoodDoc = GetThisDocument()
'      If Not bGoodDoc Then AddDocument
'   End If
'
   If bCancel = 1 Then Exit Sub
   UpdateCombos
   ExitCombos
End Sub


Private Sub cboSheet_Click()
'   bGoodDoc = GetThisDocument
   UpdateCombos
End Sub

Private Sub cboSheet_LostFocus()
'   cboSheet = CheckLen(cboSheet, 6)
'   If bCancel = 1 Then Exit Sub
'   bGoodDoc = GetThisDocument()
'   If Not bGoodDoc Then AddDocument
'
   If bCancel = 1 Then Exit Sub
   UpdateCombos
   ExitCombos
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdDockLok_Click()
   
   Dim strRevNum As String
   Dim strSheet As String
   Dim strPartNum As String
   Dim strObsolDate As String
   Dim strEffDate As String
   
   Dim mfileInt As MfileIntegrator
   Set mfileInt = New MfileIntegrator

   Dim strIniPath As String
   strIniPath = App.Path & "\" & "MFileInit.ini"

   strPartNum = GetSectionEntry("MFILES_METADATA", "PARTNUM", strIniPath)
   strRevNum = GetSectionEntry("MFILES_METADATA", "REVNUM", strIniPath)
   strSheet = GetSectionEntry("MFILES_METADATA", "SHEET", strIniPath)
   strObsolDate = GetSectionEntry("MFILES_METADATA", "OBSOLDATE", strIniPath)
   strEffDate = GetSectionEntry("MFILES_METADATA", "EFFDATE", strIniPath)

   mfileInt.OpenXMLFile "Insert", "EngDocument", "Engineering Documents", mfileInt.gsDOCClassID, "Scan", Trim(cboDoc)
   mfileInt.AddXMLMetaData "PartNumber", strPartNum, Trim(cboDoc)
   mfileInt.AddXMLMetaData "RevisionNumber", strRevNum, Trim(cboRev)
   mfileInt.AddXMLMetaData "Sheet", strSheet, Trim(cboSheet)
   mfileInt.AddXMLMetaData "ObsoleteDate", strObsolDate, txtEdte, MFDatatypeDate
   mfileInt.AddXMLMetaData "EffectiveDate", strEffDate, txtOdte, MFDatatypeDate
   
   mfileInt.CloseXMLFile
   If mfileInt.SendXMLFileToMFile Then SysMsg "Record Successfully Indexed", True

   Set mfileInt = Nothing
   
End Sub

Private Function GetSectionEntry(ByVal strSectionName As String, ByVal strEntry As String, ByVal strIniPath As String) As String
   
   Dim X As Long
   Dim sSection As String, sEntry As String, sDefault As String
   Dim sRetBuf As String, iLenBuf As Integer, sFileName As String
   Dim sValue As String

   On Error GoTo modErr1
   
   sSection = strSectionName
   sEntry = strEntry
   sDefault = ""
   sRetBuf = String(256, vbNull) '256 null characters
   iLenBuf = Len(sRetBuf)
   sFileName = strIniPath
   X = GetPrivateProfileString(sSection, sEntry, _
                     "", sRetBuf, iLenBuf, sFileName)
   sValue = Trim(Left$(sRetBuf, X))
   
   If sValue <> "" Then
      GetSectionEntry = sValue
   Else
      GetSectionEntry = ""
   End If
   
   Exit Function
   
modErr1:
   sProcName = "GetSectionEntry"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3301
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdNew_Click()
   ExitCombos
End Sub

Private Sub cmdVew_Click()
   If cmdVew Then
      DocuDCe01b.Show
      cmdVew = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
'Debug.Print "Activate class " & PassedClass & " doc " & PassedDoc & " rev " & PassedRev & " sht " & PassedSheet
      'sSql = "update DdocTable SET DOEXTDESC='' WHERE DOEXTDESC IS NULL"
      'clsADOCon.ExecuteSQL sSql ' rdExecDirect
      FillCustomers
      cmbCst = "NONE"
      AddComboStr cmbCst.hwnd, "NONE"
      txtNme = "Not Customer Specific"
      
      Set doc = New ClassDoc
      doc.FillClasses cboClass, True
      doc.FillDocuments cboClass, cboDoc
      
      'if a specific document has been passed, select it
      If PassedClass <> "" Then
         Me.txtDsc.SetFocus
         Me.cboClass = PassedClass
         DoEvents
         Me.cboDoc = PassedDoc
         DoEvents
         GetRevisions
         Me.cboRev = PassedRev
         DoEvents
         GetSheets
         Me.cboSheet = PassedSheet
         DoEvents
         'GetThisDocument
'Debug.Print "Activated.  cboClass = " & cboDoc
      End If
      
      UpdateCombos
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub


'Private Sub GetSomeClass()
'   Dim RdoCls As ADODB.Recordset
'   Dim sClass As String
'   sClass = Compress(cboClass)
'   On Error GoTo DiaErr1
'   sSql = "Qry_GetDocumentClass '" & sClass & "' "
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoCls, ES_KEYSET)
'   If bSqlRows Then
'      With RdoCls
'         cboClass = "" & Trim(!DCLNAME)
'         lblDsc = "" & Trim(!DCLDESC)
'
'         'just show sheets regardless.  even if they are not used, they will be blank
'         'If !DCLSHEETS Then
'            z1(5).Visible = True
'            cboSheet.Visible = True
'         'Else
'         '   z1(5).Visible = False
'         '   cboSheet.Visible = False
'         'End If
'         If !DCLADCN Then
'            z1(10).Visible = True
'            z1(13).Visible = True
'            txtAdc.Visible = True
'            txtAdte.Visible = True
'         Else
'            z1(10).Visible = False
'            z1(13).Visible = False
'            txtAdc.Visible = False
'            txtAdte.Visible = False
'         End If
'         ClearResultSet RdoCls
'      End With
'   Else
'      lblDsc = ""
'      'z1(5).Visible = False
'      'cboSheet.Visible = False
'   End If
'   Set RdoCls = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getsomecla"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
Private Sub Form_Load()
'Debug.Print "Load class " & PassedClass & " doc " & PassedDoc & " rev " & PassedRev & " sht " & PassedSheet
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sLastDocClass = cboClass
   SaveSetting "Esi2000", "EsiEngr", "DocClass", sLastDocClass
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)

   If (chkFromDocLst <> vbChecked) Then
      FormUnload
   End If
   On Error Resume Next
   Set RdoDoc = Nothing
   Set DocuDCe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub txtAdc_Change()
   If Not bOnLoad Then
      'If bGoodDoc And Len(txtAdc) Then
         'txtAdte = Format(ES_SYSDATE, "mm/dd/yy")
      'Else
      '   txtAdte = ""
      'End If
   End If
   
End Sub


Private Sub txtAdc_LostFocus()
   txtAdc = CheckLen(txtAdc, 20)
   'If bGoodDoc Then
      'On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DOADCN = "" & txtAdc
      If IsDate(txtAdte) Then
         RdoDoc!DODATEADCN = Format(txtAdte, "mm/dd,yyyy")
      End If
      If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoDoc.Update
      'If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtAdte_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'endif
End Sub

Private Sub txtAdte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtAdte_LostFocus()
   If Trim(txtAdte) = "" Then
      'If bGoodDoc Then
         On Error Resume Next
         'RdoDoc.Edit
         RdoDoc!DODATEADCN = Null
         RdoDoc.Update
         If Err > 0 Then ValidateEdit
      'End If
   Else
      txtAdte = CheckDateEx(txtAdte)
      'If bGoodDoc Then
         On Error Resume Next
         'RdoDoc.Edit
         RdoDoc!DODATEADCN = Format(txtAdte, "mm/dd,yyyy")
         If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
         RdoDoc.Update
         If Err > 0 Then ValidateEdit
      'End If
   End If
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 60)
   txtDsc = StrCase(txtDsc)
   'If bGoodDoc Then
      On Error Resume Next
      ''RdoDoc.Edit
      RdoDoc!DODESCR = "" & txtDsc
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtEco_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'endif
End Sub

Private Sub txtEco_LostFocus()
   txtEco = CheckLen(txtEco, 3)
   txtEco = Format(Abs(Val(txtEco)), "##0")
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DOECO = Val(txtEco)
      If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtEdte_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'end if
End Sub

Private Sub txtEdte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEdte_LostFocus()
   If Trim(txtEdte) = "" Then
      'If bGoodDoc Then
         On Error Resume Next
         'RdoDoc.Edit
         RdoDoc!DOEFFECTIVE = Null
         RdoDoc.Update
         If Err > 0 Then ValidateEdit
      'End If
   Else
      txtEdte = CheckDateEx(txtEdte)
      'If bGoodDoc Then
         On Error Resume Next
         'RdoDoc.Edit
         RdoDoc!DOEFFECTIVE = Format(txtEdte, "mm/dd,yyyy")
         If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
         RdoDoc.Update
         If Err > 0 Then ValidateEdit
      'End If
   End If
End Sub


Private Sub txtExt_LostFocus()
   txtExt = CheckLen(txtExt, 2048)
   txtExt = StrCase(txtExt, ES_FIRSTWORD)
   'If bGoodDoc Then
      On Error Resume Next
      ''RdoDoc.Edit
      RdoDoc!DOEXTDESC = "" & txtExt
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtFle_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'endif
End Sub

Private Sub txtFle_LostFocus()
   txtFle = CheckLen(txtFle, 40)
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DOFILENAME = "" & txtFle
      If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtLoc_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'endif
End Sub

Private Sub txtLoc_LostFocus()
   txtLoc = CheckLen(txtLoc, 6)
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DOLOC = "" & Trim(txtLoc)
      If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtMic_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'end if
End Sub

Private Sub txtMic_LostFocus()
   txtMic = CheckLen(txtMic, 8)
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      
      RdoDoc!DOMIC = "" & txtMic
      If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtNte_LostFocus()
   txtNte = CheckLen(txtNte, 2304)
   txtNte = StrCase(txtNte, ES_FIRSTWORD)
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DONOTES = "" & txtNte
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtOdte_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'endif
End Sub

Private Sub txtOdte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtOdte_LostFocus()
   If Trim(txtOdte) = "" Then
      'If bGoodDoc Then
         On Error Resume Next
         'RdoDoc.Edit
         RdoDoc!DOOBSOLETE = Null
         RdoDoc.Update
         If Err > 0 Then ValidateEdit
      'End If
   Else
      txtOdte = CheckDateEx(txtOdte)
      'If bGoodDoc Then
         On Error Resume Next
         'RdoDoc.Edit
         RdoDoc!DOOBSOLETE = Format(txtOdte, "mm/dd,yyyy")
         If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
         RdoDoc.Update
         If Err > 0 Then ValidateEdit
      'End If
   End If
End Sub


Private Sub txtQty_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'end if
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 3)
   txtQty = Format(Abs(Val(txtQty)), "##0")
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DOQTY = Val(txtQty)
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtRdte_Change()
   'If bGoodDoc Then
      bDataChanged = True
   'end if
End Sub

Private Sub txtRdte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtRdte_LostFocus()
   txtRdte = CheckDateEx(txtRdte)
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DODATEREC = Format(txtRdte, "mm/dd,yyyy")
      If bDataChanged Then RdoDoc!DODATEREV = Format(ES_SYSDATE, "mm/dd/yyyy")
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtSze_LostFocus()
   txtSze = CheckLen(txtSze, 1)
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DOSIZE = "" & txtSze
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub


Private Sub txtTyp_LostFocus()
   txtTyp = CheckLen(txtTyp, 30)
   'If bGoodDoc Then
      On Error Resume Next
      'RdoDoc.Edit
      RdoDoc!DOTYPE = "" & txtTyp
      RdoDoc.Update
      If Err > 0 Then ValidateEdit
   'End If
   
End Sub



Private Sub ClearBoxes()
   Dim iList As Integer
   For iList = 0 To Controls.Count - 1
      If TypeOf Controls(iList) Is TextBox Then _
                         Controls(iList).Text = ""
   Next
   txtRdte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtQty = "1"
   
End Sub

Private Sub GetRevisions()
   Dim RdoRev As ADODB.Recordset
   Dim sDocument As String
   cboSheet.Clear
   cboRev.Clear
   sDocument = Compress(cboDoc)
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT DOREF,DOREV FROM DdocTable " _
          & "WHERE DOREF='" & sDocument & "' AND DOCLASS = '" & cboClass & "'"
'Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev, ES_FORWARD)
   If bSqlRows Then
      With RdoRev
         Do Until .EOF
            cboRev.AddItem "" & Trim(!DOREV)
            .MoveNext
         Loop
         ClearResultSet RdoRev
      End With
   End If
   If cboRev.ListCount > 0 Then
      cboRev = cboRev.List(0)
      GetSheets
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getrevisi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetSheets()
   Dim sClass As String
   Dim sRevision As String
   Dim sDocument As String
   cboSheet.Clear
   sClass = Compress(cboClass)
   sDocument = Compress(cboDoc)
   sRevision = Compress(cboRev)
   
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT DOSHEET FROM DdocTable " _
          & "WHERE DOREF='" & sDocument & "' AND DOREV='" & sRevision & "' " _
          & "AND DOCLASS='" & sClass & "' "
'Debug.Print sSql
   LoadComboBox cboSheet, -1
   If cboSheet.ListCount > 0 Then cboSheet = cboSheet.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "getsheets"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetThisDocument() As Boolean
   Dim sClass As String
   Dim sDocument As String
   Dim sRevision As String
   Dim sSheet As String
   
   sClass = Compress(cboClass)
   sDocument = Compress(cboDoc)
   sRevision = Compress(cboRev)
   sSheet = Compress(cboSheet)
   'On Error Resume Next
   bDataChanged = False
   'RdoDoc.Close
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM DdocTable WHERE DOREF='" & sDocument & "' " _
          & "AND DOREV='" & sRevision & "' AND DOSHEET='" & sSheet & "' " _
          & "AND DOCLASS='" & sClass & "' "
'Debug.Print sSql & " (" & PassedDoc & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_KEYSET)
   If bSqlRows Then
      With RdoDoc
         cboDoc = "" & Trim(!DONUM)
         cboRev = "" & Trim(!DOREV)
         cboSheet = "" & Trim(!DOSHEET)
         txtDsc = "" & Trim(!DODESCR)
         txtExt = "" & Trim(!DOEXTDESC)
         txtMic = "" & Trim(!DOMIC)
         txtSze = "" & Trim(!DOSIZE)
         txtEco = "" & Format(!DOECO, "##0")
         txtLoc = "" & Trim(!DOLOC)
         txtQty = "" & Format(!DOQTY, "##0")
         txtAdc = "" & Trim(!DOADCN)
         txtTyp = "" & Trim(!DOTYPE)
         txtFle = "" & Trim(!DOFILENAME)
         txtRdte = "" & Format(!DODATEREC, "mm/dd/yyyy")
         txtAdte = "" & Format(!DODATEADCN, "mm/dd/yyyy")
         txtEdte = "" & Format(!DOEFFECTIVE, "mm/dd/yyyy")
         txtOdte = "" & Format(!DOOBSOLETE, "mm/dd/yyyy")
         txtNte = "" & Trim(!DONOTES)
         If Trim(!DOCUST) = "NONE" Then
            cmbCst = "NONE"
            txtNme = "Not Customer Specific"
         Else
            cmbCst = "" & Trim(!DOCUST)
            FindCustomer Me, cmbCst, False
         End If
         GetThisDocument = True
      End With
      sOldDoc = Trim(cboDoc)
   Else
      GetThisDocument = False
      If sOldDoc <> Trim(cboDoc) Then
         txtDsc = ""
         txtExt = ""
         txtTyp = ""
      End If
      txtMic = ""
      txtSze = ""
      txtEco = "0"
      txtLoc = ""
      txtQty = "1"
      txtAdc = ""
      txtFle = ""
      txtRdte = Format(ES_SYSDATE, "mm/dd/yyyy")
      txtAdte = ""
      txtEdte = ""
      txtOdte = ""
      txtNte = ""
   End If
   'cmdNew.Enabled = Not GetThisDocument
   Exit Function
   
DiaErr1:
   sProcName = "getthisdoc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddDocument()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sClass As String
   Dim sCust As String
   Dim sDocument As String
   Dim sRevision As String
   Dim sSheet As String
   
   bResponse = IllegalCharacters(cboDoc)
   If bResponse > 0 Then
      MsgBox "The Document Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   bResponse = IllegalCharacters(cboRev)
   If bResponse > 0 Then
      MsgBox "The Revision Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   On Error Resume Next
   If Trim(cboDoc) = "" Then
      MsgBox "Requires A Valid Document.", vbExclamation, Caption
      cboDoc.SetFocus
      Exit Sub
   End If
   'If cboSheet.Visible Then
   '   If Trim(cboSheet) = "" Then
   '      MsgBox "Requires A Valid Sheet.", vbExclamation, Caption
   '      cboSheet.SetFocus
   '      Exit Sub
   '   End If
   'End If
   On Error GoTo 0
   sClass = Compress(cboClass)
   sCust = Compress(cmbCst)
   sDocument = Compress(cboDoc)
   sRevision = Compress(cboRev)
   sSheet = Compress(cboSheet)
   'cboRev = sRevision
   'cboSheet = sSheet
   
'   sMsg = "That Document Wasn't Found." & vbCrLf _
'          & "Add class " & cboClass & " doc " & cboDoc & " rev " & sRevision & " sheet " & sSheet & "?"
'   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   
   bResponse = vbYes
   
   If bResponse = vbYes Then
'      If sOldDoc <> Trim(cboDoc) Then
'         sSql = "INSERT INTO DdocTable (DOREF,DONUM,DOREV,DOCLASS,DOSHEET,DOCUST) " _
'                & "VALUES('" & sDocument & "','" & cboDoc & "','" _
'                & sRevision & "','" & sClass & "','" & sSheet & "','" & sCust & "')"
'      Else
'         sSql = "INSERT INTO DdocTable (DOREF,DONUM,DOREV,DOCLASS," _
'                & "DOSHEET,DOCUST,DODESCR,DOEXTDESC,DOTYPE) VALUES('" _
'                & sDocument & "','" _
'                & cboDoc & "','" & sRevision & "','" & sClass & "','" _
'                & sSheet & "','" & sCust & "','" & Trim(txtDsc) & "','" _
'                & txtExt & "','" & Trim(txtTyp) & "')"
'      End If
      
      sSql = "INSERT INTO DdocTable (DOREF,DONUM,DOREV,DOCLASS,DOSHEET) " & vbCrLf _
         & "VALUES('" & sDocument & "','" & cboDoc & "','" _
         & sRevision & "','" & sClass & "','" & sSheet & "')"
      
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         
         'sSql = "update DdocTable SET DOEXTDESC='' WHERE " _
         '       & "DOEXTDESC IS NULL"
         'clsADOCon.ExecuteSQL sSql ' rdExecDirect
         If Len(sRevision) > 0 Then cboRev.AddItem cboRev
         If cboSheet.Visible Then cboSheet.AddItem cboSheet
         SysMsg "Document Added.", True
         'bGoodDoc = GetThisDocument()
         fraEdit.Enabled = GetThisDocument
      Else
         MsgBox "Couldn't Add The Document.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "adddocume"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdateCombos()

   If updatingCombos Then
      Exit Sub
   Else
      updatingCombos = True
   End If

   Dim rippleDown As Boolean
   If Compress(cboClass) <> sClass Then
      sClass = Compress(cboClass)
      rippleDown = True
      lblDsc = doc.GetClassDesc(cboClass)
      doc.FillDocuments cboClass, cboDoc
   End If
   
   If Compress(cboDoc) <> sDoc Or rippleDown Then
      sDoc = Compress(cboDoc)
      rippleDown = True
      doc.FillRevisions cboClass, cboDoc, cboRev
   End If
      
   If Compress(cboRev) <> sRev Or rippleDown Then
      sRev = Compress(cboRev)
      rippleDown = True
      doc.FillSheets cboClass, cboDoc, cboRev, cboSheet
   End If
      
   sSheet = cboSheet
   'lblDocDesc = GetDocDesc
   
   'if document exists, show it.  Otherwise, disable editing
   Me.fraEdit.Enabled = GetThisDocument
   
   updatingCombos = False
End Sub

Private Sub ExitCombos()
   'if exiting combos, verify whether new document should be created
   
   'if still loading, exit
   If bOnLoad Then
      Exit Sub
   End If
   
   'if still in top combos, no action required
   Select Case Me.ActiveControl.Name
   Case "cboClass", "cboDoc", "cboRev", "cboSheet"
      Exit Sub
   End Select
   
   'if editing the current document, no action required
   'If fraEdit.Enabled Then
   '   Exit Sub
   'End If
   
   'require a nonblank document number
   If Len(Trim(cboDoc)) = 0 Then
      MsgBox "A valid document number is required"
      Exit Sub
   End If
   
   'read the document, if any
   If GetThisDocument Then
      Exit Sub
   End If
   
   'otherwise, ask whether to create the document
   Dim msg As String
   msg = "Create class " & cboClass & " document " & cboDoc
   If cboRev.Text <> "" Then
      msg = msg & " rev " & cboRev
   End If
   If cboSheet.Text <> "" Then
      msg = msg & " sheet " & cboSheet
   End If
   msg = msg & "?"
   
   Select Case MsgBox(msg, vbYesNoCancel, "Create new document?")
   Case vbYes
   Case Else
      Exit Sub
   End Select

   AddDocument
End Sub

