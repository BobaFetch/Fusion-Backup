VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Documents To Parts"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   32
      Top             =   1440
      Width           =   7692
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtFolder 
      Height          =   288
      Left            =   1440
      TabIndex        =   17
      Tag             =   "2"
      ToolTipText     =   "Default Starting Folder For Files (Workstation Setting)"
      Top             =   4680
      Width           =   5050
   End
   Begin VB.CheckBox optHelp 
      Caption         =   "for help"
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Find The Path"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   15
      ToolTipText     =   "Full Path To Link (See Help). (50) Chars "
      Top             =   3960
      Width           =   5050
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   16
      ToolTipText     =   "Document/Link Description. (50) Chars"
      Top             =   4320
      Width           =   3615
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   11
      ToolTipText     =   "Full Path To Link (See Help). (50) Chars "
      Top             =   3240
      Width           =   5050
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   12
      ToolTipText     =   "Document/Link Description. (50) Chars"
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Find The Path"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "Full Path To Link (See Help). (50) Chars "
      Top             =   2400
      Width           =   5050
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Document/Link Description. (50) Chars"
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Find The Path"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Find The Path"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Document/Link Description. (50) Chars"
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Full Path To Link (See Help). (50) Chars "
      Top             =   1560
      Width           =   5050
   End
   Begin VB.TextBox txtDmy 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   75
   End
   Begin VB.ComboBox cmbPls 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Select Part Number (Contains Types Below Type 4)"
      Top             =   720
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6960
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   5400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5385
      FormDesignWidth =   7905
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Place PDF Documents In Assign Pictures To Parts"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   30
      Top             =   360
      Width           =   5652
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Folder"
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   29
      ToolTipText     =   "Default Starting Folder For Files (Workstation Setting)"
      Top             =   4680
      Width           =   1572
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Assignment"
      ForeColor       =   &H80000008&
      Height          =   612
      Index           =   4
      Left            =   6720
      TabIndex        =   28
      ToolTipText     =   "Double Click To Open"
      Top             =   3960
      Width           =   1116
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Assignment"
      ForeColor       =   &H80000008&
      Height          =   612
      Index           =   3
      Left            =   6720
      TabIndex        =   27
      ToolTipText     =   "Double Click To Open"
      Top             =   3240
      Width           =   1116
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Assignment"
      ForeColor       =   &H80000008&
      Height          =   612
      Index           =   2
      Left            =   6720
      TabIndex        =   26
      ToolTipText     =   "Double Click To Open"
      Top             =   2400
      Width           =   1116
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Assignment"
      ForeColor       =   &H80000008&
      Height          =   612
      Index           =   1
      Left            =   6720
      TabIndex        =   25
      ToolTipText     =   "Double Click To Open"
      Top             =   1560
      Width           =   1116
   End
   Begin VB.OLE OLE1 
      DisplayType     =   1  'Icon
      Height          =   612
      Index           =   4
      Left            =   6720
      OLETypeAllowed  =   0  'Linked
      TabIndex        =   18
      Top             =   3960
      Visible         =   0   'False
      Width           =   1104
   End
   Begin VB.OLE OLE1 
      DisplayType     =   1  'Icon
      Height          =   612
      Index           =   3
      Left            =   6720
      OLETypeAllowed  =   0  'Linked
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   1104
   End
   Begin VB.OLE OLE1 
      DisplayType     =   1  'Icon
      Height          =   612
      Index           =   2
      Left            =   6720
      OLETypeAllowed  =   0  'Linked
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1104
   End
   Begin VB.OLE OLE1 
      DisplayType     =   1  'Icon
      Height          =   612
      Index           =   1
      Left            =   6720
      OLETypeAllowed  =   0  'Linked
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1104
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   6
      Left            =   4440
      TabIndex        =   21
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblLvl 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5040
      TabIndex        =   20
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "DocuDCe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'3/2/06 Corrected closing the wrong form
'4/13/06 Added pdf warning
Option Explicit
'Dim RdoQry As rdoQuery
'Dim RdoPdc As ADODB.Recordset
Dim AdoCmdObj As ADODB.Command
Dim RdoPdc As ADODB.Recordset

Dim bCanceled As Boolean
Dim bOnLoad As Byte
Dim bGoodPart As Byte
Dim bIndex As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbPls_Click()
   bGoodPart = GetThisPart()
   
End Sub


Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   
   If (Not ValidPartNumber(cmbPls.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPls = ""
      Exit Sub
   End If
   
   If bCanceled = False Then bGoodPart = GetThisPart()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3304
      MouseCursor 0
      cmdHlp.Value = False
   End If
   bCanceled = False
   
End Sub

Private Sub cmdHlp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdSel_Click(Index As Integer)
   bIndex = Index
   If bGoodPart = 1 Then
      On Error Resume Next
      MDISect.Cdi.InitDir = txtFolder
      MDISect.Cdi.DialogTitle = "Show Files"
      MDISect.Cdi.Filter = "Document Files (*.doc)|*.doc|" _
                           & "Excel Files(*.xls)|*.xls|" _
                           & "HTML Files(*.html)|*.html|" _
                           & "HTM Files(*.htm)|*.htm|" _
                           & "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
      MDISect.Cdi.FilterIndex = 1
      MDISect.Cdi.ShowOpen
      If Trim(MDISect.Cdi.FileName) <> "" Then txtPath(Index) = MDISect.Cdi.FileName
      With RdoPdc
         '.Edit
         Select Case Index
            Case 1
               !PADOCLINK1 = Trim(txtPath(Index))
            Case 2
               !PADOCLINK2 = Trim(txtPath(Index))
            Case 3
               !PADOCLINK3 = Trim(txtPath(Index))
            Case Else
               !PADOCLINK4 = Trim(txtPath(Index))
         End Select
         .Update
      End With
   Else
      MsgBox "Requires A Valid Part Number.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillPartsBelow4 cmbPls
      If cmbPls.ListCount > 0 Then
         cmbPls = cmbPls.List(0)
         bGoodPart = GetThisPart()
      End If
      bOnLoad = 0
   End If
   If optHelp.Value = vbChecked Then Unload HelpLink
   optHelp.Value = vbUnchecked
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   txtFolder = GetSetting("Esi2000", "EsiEngr", "docpath", txtFolder)
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PADOCLINK1,PADOCDESC1," _
          & "PADOCLINK2,PADOCDESC2,PADOCLINK3,PADOCDESC3," _
          & "PADOCLINK4,PADOCDESC4 FROM PartTable WHERE PARTREF = ?"
   
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmPrtRef As ADODB.Parameter
   Set prmPrtRef = New ADODB.Parameter
   prmPrtRef.Type = adChar
   prmPrtRef.Size = 30
   AdoCmdObj.Parameters.Append prmPrtRef
   
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoCmdObj = Nothing
   Set RdoPdc = Nothing
   Set DocuDCe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDmy.BackColor = Es_FormBackColor
   txtDmy.ForeColor = Es_FormBackColor
   txtDmy = ""
   
End Sub


Private Function GetThisPart() As Byte
   On Error GoTo DiaErr1
   AdoCmdObj.Parameters(0).Value = Compress(cmbPls)
   bSqlRows = clsADOCon.GetQuerySet(RdoPdc, AdoCmdObj, ES_KEYSET, True, 1)
   If bSqlRows Then
      MouseCursor 13
      With RdoPdc
         On Error GoTo DiaErr1
         cmbPls = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = Format(!PALEVEL, "0")
         On Error Resume Next
         txtPath(1) = "" & Trim(!PADOCLINK1)
         If Len(Trim(!PADOCLINK1)) > 0 Then
            Box(1).Visible = False
            OLE1(1).DisplayType = 1
            OLE1(1).CreateLink txtPath(1)
            OLE1(1).Visible = True
         Else
            Box(1).Visible = True
            OLE1(1).DisplayType = 0
            OLE1(1).Visible = False
         End If
         txtDesc(1) = "" & Trim(!PADOCDESC1)
         
         txtPath(2) = "" & Trim(!PADOCLINK2)
         If Len(Trim(!PADOCLINK2)) > 0 Then
            Box(2).Visible = False
            OLE1(2).DisplayType = 1
            OLE1(2).CreateLink txtPath(2)
            OLE1(2).Visible = True
         Else
            Box(2).Visible = True
            OLE1(2).DisplayType = 0
            OLE1(2).Visible = False
         End If
         txtDesc(2) = "" & Trim(!PADOCDESC2)
         
         txtPath(3) = "" & Trim(!PADOCLINK3)
         If Len(Trim(!PADOCLINK3)) > 0 Then
            Box(3).Visible = False
            OLE1(3).DisplayType = 1
            OLE1(3).CreateLink txtPath(3)
            OLE1(3).Visible = True
         Else
            Box(3).Visible = True
            OLE1(3).DisplayType = 0
            OLE1(3).Visible = False
         End If
         txtDesc(3) = "" & Trim(!PADOCDESC3)
         
         txtPath(4) = "" & Trim(!PADOCLINK4)
         If Len(Trim(!PADOCLINK4)) > 0 Then
            Box(4).Visible = False
            OLE1(4).DisplayType = 1
            OLE1(4).CreateLink txtPath(4)
            OLE1(4).Visible = True
         Else
            Box(4).Visible = True
            OLE1(4).DisplayType = 0
            OLE1(4).Visible = False
         End If
         txtDesc(4) = "" & Trim(!PADOCDESC4)
         
         txtDmy.Enabled = False
         GetThisPart = 1
      End With
   Else
      lblDsc = "*** Part Number Wasn't Found."
      lblLvl = ""
      txtDmy.Enabled = True
      GetThisPart = 0
   End If
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getthispart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub txtDesc_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtDesc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtDesc_Validate(Index As Integer, Cancel As Boolean)
   txtDesc(Index) = CheckLen(txtDesc(Index), 50)
   txtDesc(Index) = StrCase(txtDesc(Index))
   On Error Resume Next
   If bGoodPart = 1 Then
      'RdoPdc.Edit
      Select Case Index
         Case 1
            RdoPdc!PADOCDESC1 = txtDesc(Index)
         Case 2
            RdoPdc!PADOCDESC2 = txtDesc(Index)
         Case 3
            RdoPdc!PADOCDESC3 = txtDesc(Index)
         Case Else
            RdoPdc!PADOCDESC4 = txtDesc(Index)
      End Select
      RdoPdc.Update
   End If
   
End Sub


Private Sub txtFolder_LostFocus()
   SaveSetting "Esi2000", "EsiEngr", "docpath", txtFolder
   
End Sub


Private Sub txtPath_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtPath_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub

Private Sub txtPath_LostFocus(Index As Integer)
   On Error GoTo DiaErr1:
   bIndex = Index
   If LCase(Right(txtPath(Index), 4)) = ".pdf" Then
      MsgBox "Please Place .pdf Files In Assign Pictures.", _
         vbInformation, Caption
      Exit Sub
   End If
   txtPath(Index) = CheckLen(txtPath(Index), 80)
   If Len(txtPath(Index)) > 0 Then
      Box(Index).Visible = False
      OLE1(Index).DisplayType = 1
      OLE1(Index).Visible = True
      OLE1(Index).CreateLink txtPath(Index)
   Else
      Box(Index).Visible = True
      OLE1(Index).DisplayType = 0
      OLE1(Index).Visible = False
   End If
   On Error Resume Next
   If bGoodPart = 1 Then
      'RdoPdc.Edit
      Select Case Index
         Case 1
            RdoPdc!PADOCLINK1 = Trim(txtPath(Index))
         Case 2
            RdoPdc!PADOCLINK2 = Trim(txtPath(Index))
         Case 3
            RdoPdc!PADOCLINK3 = Trim(txtPath(Index))
         Case Else
            RdoPdc!PADOCLINK4 = Trim(txtPath(Index))
      End Select
      RdoPdc.Update
   End If
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   MsgBox "Selection Wasn't Found Or Does Not Support" _
      & "Object Linking And Embedding.", vbInformation, _
      Caption
   OLE1(Index).Visible = False
   
End Sub
