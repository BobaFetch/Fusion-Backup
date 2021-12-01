VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form DocuDCe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Pictures To Parts"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Left            =   180
      TabIndex        =   38
      Top             =   1420
      Width           =   7692
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCe05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtFolder 
      Height          =   288
      Left            =   1440
      TabIndex        =   14
      Tag             =   "2"
      ToolTipText     =   "Default Starting Folder For Files (Workstation Setting)"
      Top             =   4800
      Width           =   4812
   End
   Begin VB.CheckBox optVew 
      Caption         =   "for viewers"
      Height          =   255
      Left            =   3840
      TabIndex        =   32
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdVew 
      Caption         =   "S&ettings"
      Height          =   300
      Left            =   6960
      TabIndex        =   31
      ToolTipText     =   "Change Viewers/Locations For This Workstation"
      Top             =   600
      Width           =   875
   End
   Begin VB.CheckBox optHelp 
      Caption         =   "for help"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Find The Path"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   12
      ToolTipText     =   "Full Path To Link (See Help). (80) Chars Max "
      Top             =   3960
      Width           =   5050
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   13
      ToolTipText     =   "Document/Link Description. (50) Chars"
      Top             =   4320
      Width           =   3615
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Full Path To Link (See Help). (80) Chars Max "
      Top             =   3240
      Width           =   5050
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   10
      ToolTipText     =   "Document/Link Description. (50) Chars"
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Find The Path"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      ToolTipText     =   "Full Path To Link (See Help). (80) Chars Max "
      Top             =   2400
      Width           =   5050
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "Document/Link Description. (50) Chars"
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   5
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
      Width           =   3612
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Full Path To Link (See Help). (80) Chars Max "
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
      Height          =   288
      Left            =   1440
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   5640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5595
      FormDesignWidth =   7965
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Folder"
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   36
      ToolTipText     =   "Default Starting Folder For Files (Workstation Setting)"
      Top             =   4800
      Width           =   1572
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(PDF And JPGs Use Windows Default)"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   4680
      TabIndex        =   35
      Top             =   5640
      Visible         =   0   'False
      Width           =   3804
   End
   Begin VB.Label lblWeb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Windows Browser"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1080
      TabIndex        =   34
      Top             =   6360
      Visible         =   0   'False
      Width           =   5004
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Web"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   33
      Top             =   6360
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   30
      Top             =   6120
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Media"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pictures"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label lblTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1080
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   5004
   End
   Begin VB.Label lblMMV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1080
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   5004
   End
   Begin VB.Label lblPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1080
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   3804
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No File Assigned"
      ForeColor       =   &H80000008&
      Height          =   612
      Index           =   4
      Left            =   6720
      TabIndex        =   24
      ToolTipText     =   "Dbl Click To View Picture"
      Top             =   3960
      Width           =   1116
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No File Assigned"
      ForeColor       =   &H80000008&
      Height          =   612
      Index           =   3
      Left            =   6720
      TabIndex        =   23
      ToolTipText     =   "Dbl Click To View Picture"
      Top             =   3240
      Width           =   1116
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No File Assigned"
      ForeColor       =   &H80000008&
      Height          =   612
      Index           =   2
      Left            =   6720
      TabIndex        =   22
      ToolTipText     =   "Dbl Click To View Picture"
      Top             =   2400
      Width           =   1116
   End
   Begin VB.Label Box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No File Assigned"
      ForeColor       =   &H80000008&
      Height          =   612
      Index           =   1
      Left            =   6720
      TabIndex        =   21
      ToolTipText     =   "Dbl Click To View Picture"
      Top             =   1560
      Width           =   1116
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1440
      TabIndex        =   19
      Top             =   1080
      Width           =   3132
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   252
      Index           =   6
      Left            =   4680
      TabIndex        =   17
      Top             =   1080
      Width           =   612
   End
   Begin VB.Label lblLvl 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   5280
      TabIndex        =   16
      Top             =   1080
      Width           =   372
   End
End
Attribute VB_Name = "DocuDCe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/2/05 Changed the path and picture to (80) chars
'12/2/05 Added RefreshData to correct erroneous character sent by _
'        the TextBox
'1/3/06 Added PDF and html files to default
'3/2/06 Corrected closing the wrong form
'3/9/06 Changed cmdSel(index) to update rows and update Box(Index)
Option Explicit
'Dim RdoQry As rdoQuery
'Dim RdoPdc As ADODB.Recordset

Dim AdoCmdObj As ADODB.Command
Dim RdoPdc As ADODB.Recordset

Dim bCanceled As Boolean
Dim bOnLoad As Byte
Dim bGoodPart As Byte
Dim bLocalIndex As Byte

Dim sPicViewer As String
Dim sMMViewer As String
Dim sTxtViewer As String
Dim sWebViewer As String

Dim sPictures(5) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub Box_DblClick(Index As Integer)
   Dim bByte As Byte
   bLocalIndex = Index
   RefreshData
   OpenThisPicture bLocalIndex
   
End Sub


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
   
   If Not bCanceled Then bGoodPart = GetThisPart()
   
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
      OpenHelpContext 3305
      MouseCursor 0
      cmdHlp.Value = False
   End If
   
End Sub


Private Sub cmdSel_Click(Index As Integer)
   bLocalIndex = Index
   If bGoodPart = 1 Then
      On Error Resume Next
      MDISect.Cdi.InitDir = txtFolder
      MDISect.Cdi.DialogTitle = "Show Files"
      MDISect.Cdi.Filter = "All Files (*.*)|*.*|PDF Files (*.pdf)|*.pdf|Picture Files (*.jpg)|*.jpg|" _
                           & "Picture Files(*.gif)|*.gif|" _
                           & "Picture Files(*.bmp)|*.bmp|"
      MDISect.Cdi.FilterIndex = 1
      MDISect.Cdi.ShowOpen
      If Trim(MDISect.Cdi.FileName) <> "" Then txtPath(Index) = MDISect.Cdi.FileName
      If Len(txtPath(Index)) Then
         Box(Index) = "File Assigned"
         Box(Index).ToolTipText = "Double Click To View"
      Else
         Box(Index) = "No File Assigned"
      End If
      With RdoPdc
         '.Edit
         Select Case Index
            Case 1
               !PAPICLINK1 = Trim(txtPath(Index))
            Case 2
               !PAPICLINK2 = Trim(txtPath(Index))
            Case 3
               !PAPICLINK3 = Trim(txtPath(Index))
            Case Else
               !PAPICLINK4 = Trim(txtPath(Index))
         End Select
         .Update
      End With
   Else
      MsgBox "Requires A Valid Part Number.", _
         vbInformation, Caption
   End If
   
End Sub


Private Sub cmdVew_Click()
   optVew.Value = vbChecked
   DocuViewer.txtPic = lblPic
   DocuViewer.txtMMv = lblMMV
   DocuViewer.txtTxt = lblTxt
   DocuViewer.txtWeb = lblWeb
   DocuViewer.Show
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillPartsBelow4 cmbPls
      If cmbPls.ListCount > 0 Then
         GetPictureViewers
         cmbPls = cmbPls.List(0)
         bGoodPart = GetThisPart()
      End If
      bOnLoad = 0
   End If
   If optVew.Value = vbChecked Then
      Unload DocuViewer
      optVew.Value = vbChecked
   End If
   If optHelp.Value = vbChecked Then Unload HelpLink
   optHelp.Value = vbUnchecked
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   txtFolder = GetSetting("Esi2000", "EsiEngr", "picpath", txtFolder)
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAPICLINK1,PAPICDESC1," _
          & "PAPICLINK2,PAPICDESC2,PAPICLINK3,PAPICDESC3," _
          & "PAPICLINK4,PAPICDESC4 FROM PartTable WHERE PARTREF = ?"
   
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Esi2000", "EsiEngr", "PicViewer", lblPic
   SaveSetting "Esi2000", "EsiEngr", "MMViewer", lblMMV
   SaveSetting "Esi2000", "EsiEngr", "TXTViewer", lblTxt
   SaveSetting "Esi2000", "EsiEngr", "WEBViewer", lblWeb
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoCmdObj = Nothing
   Set RdoPdc = Nothing
   Set DocuDCe05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDmy.BackColor = Es_FormBackColor
   txtDmy.ForeColor = Es_FormBackColor
   txtDmy = ""
   For b = 1 To 4
      txtPath(b).ToolTipText = "Full Path To Link (See Help). (80) Chars"
   Next
   
End Sub


Private Function GetThisPart() As Byte
   On Error GoTo DiaErr1
   Erase sPictures
   AdoCmdObj.Parameters(0).Value = Compress(cmbPls)
   bSqlRows = clsADOCon.GetQuerySet(RdoPdc, AdoCmdObj, ES_KEYSET, True, 0)
   If bSqlRows Then
      MouseCursor 13
      With RdoPdc
         On Error GoTo DiaErr1
         cmbPls = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = Format(!PALEVEL, "0")
         On Error Resume Next
         txtPath(1) = "" & Trim(!PAPICLINK1)
         If Len(Trim(!PAPICLINK1)) > 0 Then
            Box(1) = "File Assigned"
            Box(1).ToolTipText = "Double Click To View"
         Else
            Box(1) = "No File Assigned"
            Box(1).ToolTipText = "No File To View"
         End If
         txtDesc(1) = "" & Trim(!PAPICDESC1)
         
         txtPath(2) = "" & Trim(!PAPICLINK2)
         If Len(Trim(!PAPICLINK2)) > 0 Then
            Box(2) = "File Assigned"
            Box(2).ToolTipText = "Double Click To View"
         Else
            Box(2) = "No File Assigned"
            Box(2).ToolTipText = "No File To View"
         End If
         txtDesc(2) = "" & Trim(!PAPICDESC2)
         
         txtPath(3) = "" & Trim(!PAPICLINK3)
         If Len(Trim(!PAPICLINK3)) > 0 Then
            Box(3) = "File Assigned"
            Box(3).ToolTipText = "Double Click To View"
         Else
            Box(3) = "No File Assigned"
            Box(3).ToolTipText = "No File To View"
         End If
         txtDesc(3) = "" & Trim(!PAPICDESC3)
         
         txtPath(4) = "" & Trim(!PAPICLINK4)
         If Len(Trim(!PAPICLINK4)) > 0 Then
            Box(4) = "File Assigned"
            Box(4).ToolTipText = "Double Click To View"
         Else
            Box(4) = "No File Assigned"
            Box(4).ToolTipText = "No File To View"
         End If
         txtDesc(4) = "" & Trim(!PAPICDESC4)
         
         sPictures(1) = txtPath(1)
         sPictures(2) = txtPath(2)
         sPictures(3) = txtPath(3)
         sPictures(4) = txtPath(4)
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


Private Sub txtDesc_LostFocus(Index As Integer)
   txtDesc(Index) = CheckLen(txtDesc(Index), 50)
   txtDesc(Index) = StrCase(txtDesc(Index))
   On Error Resume Next
   If bGoodPart = 1 Then
      'RdoPdc.Edit
      Select Case Index
         Case 1
            RdoPdc!PAPICDESC1 = txtDesc(Index)
         Case 2
            RdoPdc!PAPICDESC2 = txtDesc(Index)
         Case 3
            RdoPdc!PAPICDESC3 = txtDesc(Index)
         Case Else
            RdoPdc!PAPICDESC4 = txtDesc(Index)
      End Select
      RdoPdc.Update
   End If
   
End Sub

Private Sub txtFolder_LostFocus()
   SaveSetting "Esi2000", "EsiEngr", "picpath", txtFolder
   
   
End Sub


Private Sub txtPath_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtPath_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub

Private Sub txtPath_LostFocus(Index As Integer)
   On Error GoTo DiaErr1:
   bLocalIndex = Index
   txtPath(Index) = CheckLen(txtPath(Index), 80)
   On Error Resume Next
   If bGoodPart = 1 Then
      With RdoPdc
         '.Edit
         Select Case Index
            Case 1
               !PAPICLINK1 = Trim(txtPath(Index))
            Case 2
               !PAPICLINK2 = Trim(txtPath(Index))
            Case 3
               !PAPICLINK3 = Trim(txtPath(Index))
            Case Else
               !PAPICLINK4 = Trim(txtPath(Index))
         End Select
         .Update
      End With
      If Len(txtPath(Index)) Then
         Box(Index) = "File Assigned"
         Box(Index).ToolTipText = "Double Click To View"
      Else
         Box(Index) = "No File Assigned"
         Box(Index).ToolTipText = "No File To View"
      End If
   End If
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   MsgBox "Selection Wasn't Found Or Does Not Support" _
      & "Object Viewing.", vbInformation, _
      Caption
   
End Sub

Private Sub OpenThisPicture(bIndex As Byte)
   Dim bByte As Byte
   Dim bPos As Byte
   Dim sPicType As String
   Dim sViewer As String
   Dim sWebStr As String
   Dim vRetVal As Variant
   
   sMMViewer = Trim(lblMMV)
   sPicViewer = Trim(lblPic)
   sTxtViewer = Trim(lblTxt)
   sWebViewer = Trim(lblWeb)
   
   On Error GoTo DiaErr1
   sWebStr = txtPath(bIndex)
   
   If Len(Trim(txtPath(bIndex))) > 0 Then
      sPicType = Right$(Trim$(txtPath(bIndex)), 5)
      bPos = InStr(1, sPicType, ".")
      If bPos > 0 Then
         sPicType = Right$(sPicType, 5 - bPos)
         Select Case LCase$(sPicType)
            Case "jpg", "jpeg", "pdf", "htm", "html"
               OpenWebPage sPictures(bLocalIndex)
               bByte = 1
            Case "mpg", "mpeg"
               sViewer = sMMViewer
            Case "txt"
               sViewer = sTxtViewer
            Case Else
               sViewer = sPicViewer
         End Select
         If bByte = 0 Then vRetVal = Shell(sViewer & " " & sPictures(bIndex), vbNormalFocus)
      Else
         MsgBox "Invalid Picture Format.", vbInformation, Caption
      End If
   Else
      MsgBox "No File Assigned.", vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   MsgBox "File Or Viewer Wasn't Found." & vbCrLf _
      & "Check Name And Path Of Both.", vbExclamation, _
      Caption
   
End Sub

Private Sub GetPictureViewers()
   sPicViewer = GetSetting("Esi2000", "EsiEngr", "PicViewer", sPicViewer)
   sMMViewer = GetSetting("Esi2000", "EsiEngr", "MMViewer", sMMViewer)
   sTxtViewer = GetSetting("Esi2000", "EsiEngr", "TXTViewer", sTxtViewer)
   sWebViewer = GetSetting("Esi2000", "EsiEngr", "WEBViewer", sWebViewer)
   If sPicViewer = "" Then sPicViewer = "c:\windows\system32\mspaint.exe"
   If sMMViewer = "" Then sMMViewer = "c:\Program Files\Windows Media Player\wmplayer.exe"
   If sTxtViewer = "" Then sTxtViewer = "c:\windows\system32\notepad.exe"
   If sWebViewer = "" Then sWebViewer = "c:\Program Files\Internet Explorer\IEXPLORE.EXE"
   lblPic = sPicViewer
   lblMMV = sMMViewer
   lblTxt = sTxtViewer
   'lblWeb = sWebViewer
   
End Sub

Private Sub RefreshData()
   On Error GoTo DiaErr1
   Erase sPictures
   AdoCmdObj.Parameters(0).Value = Compress(cmbPls)
   bSqlRows = clsADOCon.GetQuerySet(RdoPdc, AdoCmdObj, ES_KEYSET, True)
   If bSqlRows Then
      MouseCursor 13
      With RdoPdc
         On Error GoTo DiaErr1
         txtDesc(1) = "" & Trim(!PAPICDESC1)
         
         txtPath(2) = "" & Trim(!PAPICLINK2)
         If Len(Trim(!PAPICLINK2)) > 0 Then
            Box(2) = "File Assigned"
         Else
            Box(2) = "No File Assigned"
         End If
         txtDesc(2) = "" & Trim(!PAPICDESC2)
         
         txtPath(3) = "" & Trim(!PAPICLINK3)
         If Len(Trim(!PAPICLINK3)) > 0 Then
            Box(3) = "File Assigned"
         Else
            Box(3) = "No File Assigned"
         End If
         txtDesc(3) = "" & Trim(!PAPICDESC3)
         
         txtPath(4) = "" & Trim(!PAPICLINK4)
         If Len(Trim(!PAPICLINK4)) > 0 Then
            Box(4) = "File Assigned"
         Else
            Box(4) = "No File Assigned"
         End If
         txtDesc(4) = "" & Trim(!PAPICDESC4)
         
         sPictures(1) = Trim(!PAPICLINK1)
         sPictures(2) = Trim(!PAPICLINK2)
         sPictures(3) = Trim(!PAPICLINK3)
         sPictures(4) = Trim(!PAPICLINK4)
      End With
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "refreshcursor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
