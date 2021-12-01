VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form diaAPe05b 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invoice"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   2640
      TabIndex        =   43
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.ComboBox txtAct1 
      Height          =   315
      Index           =   5
      Left            =   3720
      TabIndex        =   4
      Tag             =   "3"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox txtAct1 
      Height          =   315
      Index           =   4
      Left            =   3720
      TabIndex        =   3
      Tag             =   "3"
      Top             =   3340
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox txtAct1 
      Height          =   315
      Index           =   3
      Left            =   3720
      TabIndex        =   2
      Tag             =   "3"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox txtAct1 
      Height          =   315
      Index           =   2
      Left            =   3720
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox txtAct1 
      Height          =   315
      Index           =   1
      Left            =   3720
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Height          =   315
      Left            =   5520
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Update List"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaAPe05b.frx":0000
      PictureDn       =   "diaAPe05b.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1680
      Top             =   4920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5610
      FormDesignWidth =   6465
   End
   Begin Threed.SSCommand cmdDn 
      Height          =   375
      Left            =   6000
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   5160
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaAPe05b.frx":028C
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   375
      Left            =   6000
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   4800
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "diaAPe05b.frx":078E
   End
   Begin VB.Label lblAct1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3720
      TabIndex        =   42
      Top             =   4440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblAct1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   41
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblAct1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3720
      TabIndex        =   40
      Top             =   3000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblAct1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   39
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblAct1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   840
      Picture         =   "diaAPe05b.frx":0C90
      Top             =   5025
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   120
      Picture         =   "diaAPe05b.frx":1182
      Top             =   5025
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   360
      Picture         =   "diaAPe05b.frx":1674
      Top             =   5025
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   600
      Picture         =   "diaAPe05b.frx":1B66
      Top             =   5025
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblPge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   5025
      Width           =   375
   End
   Begin VB.Label lblLst 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   36
      Top             =   5025
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   255
      Index           =   9
      Left            =   3960
      TabIndex        =   35
      Top             =   5025
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Index           =   10
      Left            =   5040
      TabIndex        =   34
      Top             =   5025
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account                                               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   30
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO/Notes                                                             "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   29
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   28
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   27
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblNte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblNte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblNte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblNte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblNte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   15
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblPon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   14
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblItm 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "diaAPe05b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'*************************************************************************************
' diaAPe05b - Change AP Invoice GL Distributions (items)
'
' Notes:
'
' Created: 11/13/02 (nth)
' Revisons:
'   09/30/03 (nth) Fixed accounts not updating accounts in journal
'
'**************************************************************************************

Option Explicit
Dim bOnLoad As Byte

Dim iCurrPage As Integer
Dim iIndex As Integer
Dim iTotalItems As Integer

Dim sAccount As String
Dim vInvoice(500, 5) As Variant
'0 = Invoice Item
'1 = PO
'2 = Notes
'3 = Account
'4 = Old Account

Public Sub GetAccount(iIndex)
   Dim RdoGlm As ADODB.Recordset
   Dim sCmbAccount As String
   On Error GoTo DiaErr1
   sCmbAccount = Compress(txtAct1(iIndex))
   sSql = "SELECT GLACCTREF,GLACCTNO,GLDESCR FROM GlacTable " _
          & "WHERE GLACCTREF='" & sCmbAccount & "' AND GLINACTIVE=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      txtAct1(iIndex) = "" & Trim(RdoGlm!GLACCTNO)
      lblAct1(iIndex).ForeColor = Me.ForeColor
      lblAct1(iIndex) = "" & Trim(RdoGlm!GLDESCR)
   Else
      lblAct1(iIndex).ForeColor = ES_RED
      lblAct1(iIndex) = "*** Account Wasn't Found Or Inactive ***"
   End If
   Set RdoGlm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getaccount"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   If iCurrPage > Val(lblLst) - 1 Then
      iCurrPage = Val(lblLst)
      cmdDn.enabled = False
      cmdDn.Picture = Dsdn
   Else
      cmdDn.enabled = True
      cmdDn.Picture = Endn
   End If
   If iCurrPage > 1 Then
      cmdUp.enabled = True
      cmdUp.Picture = Enup
   Else
      cmdUp.enabled = False
      cmdUp.Picture = Dsup
   End If
   lblPge = iCurrPage
   GetNextGroup
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Change AP Invoice GL Distribution"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   If iCurrPage < 2 Then
      iCurrPage = 1
      cmdUp.enabled = False
      cmdUp.Picture = Dsup
   Else
      cmdUp.enabled = True
      cmdUp.Picture = Enup
   End If
   If iCurrPage < Val(lblLst) Then
      cmdDn.enabled = True
      cmdDn.Picture = Endn
   Else
      cmdDn.enabled = False
      cmdDn.Picture = Dsdn
   End If
   lblPge = iCurrPage
   GetNextGroup
End Sub

Private Sub cmdUpd_Click()
   UpdateAccounts
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillAccounts
      GetItems
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Move 400, 400
   sCurrForm = diaAPe05a.Caption
   lblVnd = diaAPe05a.cmbVnd
   lblInv = diaAPe05a.cmbInv
   sAccount = GetSetting("Esi2000", "Fina", "LastAccount", sAccount)
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   diaAPe05a.optFrm.Value = vbUnchecked
   SaveSetting "Esi2000", "Fina", "LastAccount", sAccount
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set diaAPe05b = Nothing
End Sub

Private Sub lblAct1_Change(Index As Integer)
   If Left(lblAct1(Index), 7) = "*** Acc" Then
      lblAct1(Index).ForeColor = ES_RED
   Else
      lblAct1(Index).ForeColor = Me.ForeColor
   End If
   
End Sub

Private Sub txtAct1_Click(Index As Integer)
   GetAccount (Index)
End Sub

Private Sub txtAct1_GotFocus(Index As Integer)
   SelectFormat Me
End Sub

Private Sub txtAct1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
End Sub

Private Sub txtAct1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
End Sub

Private Sub txtAct1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
End Sub

Private Sub txtAct1_LostFocus(Index As Integer)
   txtAct1(Index) = CheckLen(txtAct1(Index), 12)
   GetAccount (Index)
   vInvoice(iIndex + Index, 3) = Compress(txtAct1(Index))
   If Len(txtAct1(Index)) Then sAccount = txtAct1(Index)
End Sub

Public Sub GetItems()
   Dim RdoItm As ADODB.Recordset
   Dim i As Integer
   Dim sVendor As String
   
   On Error GoTo DiaErr1
   sVendor = Compress(lblVnd)
   sSql = "SELECT VITNO,VITVENDOR,VITITEM,VITPO,VITPORELEASE,VITPOITEM," _
          & "VITPOITEMREV,VITACCOUNT,VITNOTE,VINO,VIVENDOR,VIDATE FROM " _
          & "ViitTable,VihdTable WHERE (VITVENDOR=VIVENDOR AND VINO=VITNO) AND " _
          & "(VINO='" & Trim(lblInv) & "' AND VIVENDOR='" & sVendor & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm)
   If bSqlRows Then
      With RdoItm
         lblDte = Format(!VIDATE, "mm/dd/yy")
         Do Until .EOF
            i = i + 1
            vInvoice(i, 0) = Format(!VITITEM, "##0")
            If !VITPO > 0 Then
               vInvoice(i, 1) = Format(!VITPO, "000000") & "-" _
                        & Trim(str(!VITPORELEASE)) & " Item " & Trim(str(!VITPOITEM)) _
                        & Trim(!VITPOITEMREV)
            Else
               vInvoice(i, 1) = "No Purchase Order"
            End If
            vInvoice(i, 2) = "" & Trim(!VITNOTE)
            vInvoice(i, 3) = "" & Trim(!VITACCOUNT)
            vInvoice(i, 4) = vInvoice(i, 3)
            .MoveNext
         Loop
      End With
   End If
   Set RdoItm = Nothing
   iTotalItems = i
   For i = 1 To iTotalItems
      If i > 5 Then Exit For
      lblItm(i).Visible = True
      lblPon(i).Visible = True
      lblNte(i).Visible = True
      txtAct1(i).Visible = True
      lblAct1(i).Visible = True
      lblItm(i) = vInvoice(i, 0)
      lblPon(i) = vInvoice(i, 1)
      lblNte(i) = vInvoice(i, 2)
      txtAct1(i) = vInvoice(i, 3)
      GetAccount (i)
   Next
   iCurrPage = 1
   lblLst = CInt((((iTotalItems * 100) / 100) / 5) + 0.5)
   If Val(lblLst) > 1 Then
      cmdDn.enabled = True
      cmdDn.Picture = Endn
   End If
   iIndex = 0
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub GetNextGroup()
   Dim i As Integer
   On Error GoTo DiaErr1
   HideBoxes
   iIndex = (Val(lblPge) - 1) * 5
   For i = 1 To 5
      If i + iIndex > iTotalItems Then Exit For
      lblItm(i).Visible = True
      lblPon(i).Visible = True
      lblNte(i).Visible = True
      txtAct1(i).Visible = True
      lblAct1(i).Visible = True
      lblItm(i) = vInvoice(i + iIndex, 0)
      lblPon(i) = vInvoice(i + iIndex, 1)
      lblNte(i) = vInvoice(i + iIndex, 2)
      txtAct1(i) = vInvoice(i + iIndex, 3)
   Next
   On Error Resume Next
   If txtAct1(1).Visible Then txtAct1(1).SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "getnextgr"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub HideBoxes()
   Dim i As Integer
   For i = 1 To 5
      lblItm(i).Visible = False
      lblNte(i).Visible = False
      lblPon(i).Visible = False
      txtAct1(i).Visible = False
      lblAct1(i).Visible = False
   Next
   
End Sub

Public Sub UpdateAccounts()
   Dim i As Integer
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sVendor As String
   Dim rdoJrn As ADODB.Recordset
   Dim iIndex As Integer
   
   On Error GoTo DiaErr1
   sVendor = Compress(lblVnd)
   sMsg = "Are Certain That You Wish To " _
          & vbCrLf & "Update Invoice " & lblInv & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      Err = 0
      bResponse = 0
      
      iIndex = (Val(lblPge) - 1) * 5
      For i = 1 To iTotalItems
         If i + iIndex > iTotalItems Then Exit For
         If Left(lblAct1(i), 3) = "***" Then
            MsgBox "At Least One Account Is Not Valid.", _
               vbInformation, Caption
            bResponse = 1
            Exit For
         End If
      Next
      
      If bResponse = 1 Then Exit Sub
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      On Error Resume Next
      sSql = "SELECT DCACCTNO FROM JritTable,JrhdTable WHERE " _
            & "JritTable.DCHEAD = JrhdTable.MJGLJRNL AND " _
            & " MJTYPE = 'PJ'AND DCCREDIT = 0 AND " _
             & "DCVENDOR = '" & sVendor & "' AND DCVENDORINV = '" _
             & lblInv & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_KEYSET)
      If bSqlRows Then
         With rdoJrn
            For i = 1 To iTotalItems
               If Len(Trim(vInvoice(i, 3))) > 0 Then
                  If vInvoice(i, 3) <> vInvoice(i, 4) Then
                     sSql = "UPDATE ViitTable SET VITACCOUNT='" & vInvoice(i, 3) _
                            & "' WHERE VITNO='" & lblInv & "' AND VITVENDOR='" _
                            & sVendor & "' AND VITITEM=" & Val(vInvoice(i, 0)) & " "
                     clsADOCon.ExecuteSQL sSql
                     .Fields(0) = vInvoice(i, 3)
                     .Update
                     If Err > 0 Then
                        Exit For
                     End If
                     
                  End If
               End If
               .MoveNext
            Next
         End With
      End If
      
      Set rdoJrn = Nothing
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "Invoice Accounts Have Been Updated.", _
            vbInformation, Caption
         Unload Me
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Could Not Update Accounts.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
   Exit Sub
DiaErr1:
   sProcName = "updateacco"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillAccounts()
   Dim rdoAct As ADODB.Recordset
   Dim i As Integer
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         Do Until .EOF
            For i = 1 To 5
               ' txtAct1(i).AddItem "" & Trim(!GLACCTNO)
               AddComboStr txtAct1(i).hWnd, "" & Trim(!GLACCTNO)
            Next
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set rdoAct = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccounts"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
