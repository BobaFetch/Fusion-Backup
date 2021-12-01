VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form BompBMe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add A Part To A Parts List"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPickAt 
      Height          =   285
      Left            =   6480
      TabIndex        =   47
      Tag             =   "1"
      Text            =   "1"
      ToolTipText     =   "Wasted (cut off)"
      Top             =   1920
      Width           =   555
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   720
      TabIndex        =   46
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdFindPart 
      Height          =   375
      Left            =   3840
      Picture         =   "BompBMe01b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMe01b.frx":043A
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   42
      Top             =   3600
      Width           =   8175
   End
   Begin VB.ComboBox cmbRev 
      Height          =   315
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Revision (Blank For Default)"
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7560
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Add This Part To The Parts List"
      Top             =   840
      Width           =   875
   End
   Begin VB.ListBox lstAssy 
      Height          =   840
      ItemData        =   "BompBMe01b.frx":0BE8
      Left            =   6840
      List            =   "BompBMe01b.frx":0BEA
      TabIndex        =   36
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "BompBMe01b.frx":0BEC
      DownPicture     =   "BompBMe01b.frx":155E
      Height          =   350
      Left            =   6240
      Picture         =   "BompBMe01b.frx":1ED0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Standard Comments"
      Top             =   2400
      Width           =   350
   End
   Begin VB.TextBox txtMatbr 
      Height          =   285
      Left            =   4440
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "Material Burden Percentage"
      Top             =   4440
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "Total Material Costs For This Level"
      Top             =   4440
      Width           =   1035
   End
   Begin VB.TextBox txtLabOh 
      Height          =   285
      Left            =   4440
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "Factory Overhead Rate"
      Top             =   4080
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Tag             =   "1"
      ToolTipText     =   "Total Accumulated Labor Cost For This Level"
      Top             =   4080
      Width           =   1035
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Sort Sequence (Otherwise Part Number)"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Quantity Used"
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtBum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Unit of Measure for Parts List"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtCvt 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Units Conversion (Feet to Inches = 12.000)"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtAdr 
      Height          =   285
      Left            =   4740
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Wasted (cut off)"
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox txtSup 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Use for Operation Testing"
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox optPhn 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4740
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtCmt 
      Height          =   1150
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "9"
      ToolTipText     =   "Comments (2048 Chars Max)"
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   7560
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5115
      FormDesignWidth =   8490
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pick At:"
      Height          =   255
      Index           =   20
      Left            =   5760
      TabIndex        =   48
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblLvl 
      Caption         =   "PALEVEL"
      Height          =   372
      Left            =   6840
      TabIndex        =   44
      Top             =   4440
      Visible         =   0   'False
      Width           =   700
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   19
      Left            =   5640
      TabIndex        =   41
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   18
      Left            =   5640
      TabIndex        =   40
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   11
      Left            =   4440
      TabIndex        =   39
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   38
      ToolTipText     =   "Revision"
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblAssy 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   35
      ToolTipText     =   "Used On Part"
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assembly "
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   34
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type "
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
      Index           =   12
      Left            =   4440
      TabIndex        =   33
      Top             =   600
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev              "
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
      Index           =   10
      Left            =   4920
      TabIndex        =   32
      Top             =   600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Burden"
      Height          =   255
      Index           =   13
      Left            =   2640
      TabIndex        =   31
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   30
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Overhead"
      Height          =   255
      Index           =   15
      Left            =   2640
      TabIndex        =   29
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Cost"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   28
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimating Costs For This Level.  Should Not Include Lower Level Costs"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seq     "
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
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                       "
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
      Index           =   1
      Left            =   720
      TabIndex        =   25
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity       "
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
      Left            =   6000
      TabIndex        =   24
      Top             =   600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Um     "
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
      Left            =   6960
      TabIndex        =   23
      Top             =   600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convert:   "
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   22
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Units Wasted:"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   21
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Qty:"
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   20
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phantom:  "
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   17
      Top             =   1200
      Width           =   3075
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4440
      TabIndex        =   16
      Top             =   840
      Width           =   405
   End
End
Attribute VB_Name = "BompBMe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'9/1/04 omit tools
'1/25/07 Fixed lblTyp (PALEVEL). Did not show Part Type.
Option Explicit
Dim bOnLoad As Byte
Dim bChanged As Byte
Dim bGoodRev As Byte
Dim bSaved As Byte
Dim bGoodPart As Byte
Dim AdoCmdObj As ADODB.Command

Dim sOldPart As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetIndexHeader()
   Dim RdoHdr As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PARTREF='" & Compress(lblAssy) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHdr, ES_FORWARD)
   If bSqlRows Then
      With RdoHdr
         lblAssy = "" & Trim(!PartNum)
         ClearResultSet RdoHdr
      End With
   End If
   Set RdoHdr = Nothing
   
End Sub

Private Function GetRevision() As Byte
   Dim RdoRes2 As ADODB.Recordset
   Dim sCurrPart As String
'   sCurrPart = Compress(cmbPrt)
   sCurrPart = Compress(cmbPrt)
   If Trim(cmbRev) = "" Then
      GetRevision = True
      Exit Function
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV FROM BmhdTable " _
          & "WHERE BMHREF='" & sCurrPart & "' AND " _
          & "BMHREV='" & Trim(cmbRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRes2)
   If bSqlRows Then
      GetRevision = True
   Else
      cmbRev = ""
      GetRevision = False
   End If
   Set RdoRes2 = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrevis"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbPrt_Click()
    GetList
End Sub

'Private Sub cmbPrt_Click()
'   GetThisPart
'   If sOldPart <> cmbPrt Then FillBomhRev cmbPrt
'
'End Sub

'Private Sub cmbPrt_LostFocus()
'   cmbPrt = CheckLen(cmbPrt, 30)
'   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
'   If lblDsc.ForeColor <> ES_RED Then
'      GetThisPart
'      If sOldPart <> cmbPrt Then FillBomhRev cmbPrt
'      sOldPart = cmbPrt
'      bChanged = 1
'   End If
'
'End Sub

Private Sub cmbPrt_LostFocus()
   Dim iCurrPartType As Integer
   Dim sOrigPart As String
    
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
    
    cmbPrt = CheckLen(cmbPrt, 30)
    sOrigPart = cmbPrt
    cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
    cmbPrt = UCase(sOrigPart)
    
    iCurrPartType = CurrentPartType(cmbPrt)
    If PartOk(cmbPrt) = 0 Then
        lblDsc.ForeColor = ES_RED
        lblDsc = "*** Part Number is Invalid ***"
    End If
    If iCurrPartType < Val(lblLvl) Or iCurrPartType > 5 Then
        lblDsc.ForeColor = ES_RED
        lblDsc = "*** Part Number is the Wrong Type ***"
    End If
    
    If lblDsc.ForeColor <> ES_RED Then
      GetThisPart
   '      If sOldPart <> txtPrt Then FillBomhRev cmbPrt
      sOldPart = cmbPrt
      bChanged = 1
   End If
End Sub


Private Sub cmbRev_Change()
   If Len(cmbRev) > 4 Then cmbRev = Left(cmbRev, 4)
   
End Sub

Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   bGoodRev = GetRevision()
   If Not bGoodRev Then
      MsgBox "That Revision Wasn't Found.", vbInformation, _
         Caption
      cmbRev = ""
   End If
   
End Sub


Private Sub cmdAdd_Click()
   Dim b As Byte
   Dim iList As Integer
   
   If Val(txtQty) = 0 Then
      MsgBox "Requires A Valid Quantity.", _
         vbInformation, Caption
      Exit Sub
   End If
   
'   For iList = 0 To BompBMe01a.lstNodes.ListCount - 1
'      If txtPrt = BompBMe01a.lstNodes.List(iList) Then
'         b = 1
'         Exit For
'      End If
'   Next
   '    If b = 1 Then
   '        lblDsc = "*** Part Number Is In Use ***"
   '        MsgBox "The Selected Part Is Used Higher " & vbCrLf _
   '            & "And Cannot Be Used On This Assembly.", vbInformation, _
   '            Caption
   '        Exit Sub
   '    End If
   
'   For iList = 0 To cmbPrt.ListCount - 1
'      If cmbPrt = cmbPrt.List(iList) Then
'         b = 1
'         Exit For
'      End If
'   Next
   
   
'   If b = 0 Then
'      lblDsc = "*** Part Number Is The Wrong Type ***"
'      MsgBox "The Selected Part Is The Wrong Part Type " & vbCrLf _
'         & "Cannot Be Used On This Assembly.", vbInformation, _
'         Caption
'      Exit Sub
'   End If
   If lblDsc.ForeColor = ES_RED Or cmbPrt = "NONE" Then
      MsgBox "Requires A Valid Part Number.", vbInformation, _
         Caption
      Exit Sub
   End If
   b = 0
   If lstAssy.ListCount > 0 Then
      For iList = 0 To lstAssy.ListCount - 1
         If cmbPrt = lstAssy.list(iList) Then b = 1
      Next
   End If
   If b = 1 Then
      MsgBox "You May Not Use The Same Part Number Twice.", _
         vbInformation, Caption
   Else
      bSaved = 1
      AddThisPart
   End If
   
End Sub

Private Sub cmdCan_Click()
   Dim bResponse As Byte
   If bSaved = 0 And bChanged = 1 Then
      bResponse = MsgBox("Exit Without Saving Changes?..", _
                  ES_NOQUESTION, Caption)
      If bResponse = vbYes Then Unload Me
   Else
      Unload Me
   End If
   
End Sub



Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 9
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdFindPart_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.lblWhereClause = "PALEVEL BETWEEN " & lblLvl & " AND 5 AND PALEVEL>0 AND PATOOL=0"
   ViewParts.Show
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3202
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      sOldPart = ""
      cmdComments.Enabled = True
      FillList
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillPartCombo
      
'      FillParts
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim b As Byte
   FormLoad Me, ES_DONTLIST
   Move BompBMe01a.Left + 400, BompBMe01a.Top + 1200
   lblLvl = BompBMe01a.lblLvl
   FormatControls
   bOnLoad = 1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF= ? "
   
  
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmPtrRef As ADODB.Parameter
   Set prmPtrRef = New ADODB.Parameter
   prmPtrRef.Type = adChar
   prmPtrRef.Size = 30
   AdoCmdObj.Parameters.Append prmPtrRef

End Sub


Private Sub FillPartCombo()
   On Error GoTo DiaErr1
   'cmbPrt.Clear
   sSql = "SELECT PARTREF,PARTNUM,PADESC From PartTable WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0"

   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.list(0)
   Exit Sub

DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub


Private Sub GetList()
   Dim RdoPls As ADODB.Recordset
   Dim sPartNumber As String
   sPartNumber = Compress(cmbPrt)
   On Error GoTo DiaErr1
   AdoCmdObj.Parameters(0) = sPartNumber
   bSqlRows = clsADOCon.GetQuerySet(RdoPls, AdoCmdObj)
   If bSqlRows Then
      With RdoPls
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         ClearResultSet RdoPls
      End With
      bGoodPart = 1
   Else
      lblDsc = ""
      MsgBox "Part Wasn't Found or Is The Wrong Type.", vbExclamation, Caption
      bGoodPart = 0
   End If
   If bGoodPart Then
      On Error Resume Next
   End If
   Set RdoPls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sCurrForm = "Bills Of Material"
   If bSaved = 1 Then BompBMe01a.optRefresh.Value = vbChecked
   BompBMe01a.cmdQuit.Enabled = True
   BompBMe01a.cmdAdd.Enabled = True
   BompBMe01a.cmdEdit.Enabled = True
   BompBMe01a.cmdCut.Enabled = True
   BompBMe01a.cmdCut.Enabled = True
   BompBMe01a.cmdCopy.Enabled = True
   BompBMe01a.cmdDelete.Enabled = True
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set BompBMe01b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = "0.000"
   txtCvt = "0.000"
   txtAdr = "0.000"
   txtSup = "0.000"
   txtLab = "0.000"
   txtLabOh = "0.000"
   txtMat = "0.000"
   txtMatbr = "0.000"
   txtSeq = "0"
   txtBum = "EA"
   txtPickAt = "1"
End Sub

Private Sub FillList()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BMASSYPART,BMPARTREF,PARTREF,PARTNUM " _
          & "FROM BmplTable,PartTable WHERE (BMPARTREF=PARTREF " _
          & "AND BMASSYPART='" & Compress(lblAssy) & "') "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
Debug.Print "Get part list for " & Compress(lblAssy)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            lstAssy.AddItem "" & Trim(!PartNum)
Debug.Print Trim(!PartNum) & " in BOM for " & Compress(lblAssy)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   BompBMe01a.cmdQuit.Enabled = True
   BompBMe01a.cmdAdd.Enabled = True
   BompBMe01a.cmdEdit.Enabled = True
   BompBMe01a.cmdCut.Enabled = True
   BompBMe01a.cmdCut.Enabled = True
   BompBMe01a.cmdCopy.Enabled = True
   BompBMe01a.cmdDelete.Enabled = True
   sProcName = "fillList"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


'Private Sub FillParts()
'   Dim RdoPrt As ADODB.Recordset
'   Dim iParts As Integer
'   Dim iNodes As Integer
'   MouseCursor 11
'   On Error GoTo DiaErr1
'   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PATOOL FROM PartTable " _
'          & "WHERE (PALEVEL BETWEEN " & lblLvl & " AND 5 " _
'          & "AND PALEVEL>0 AND PATOOL=0) ORDER BY PARTREF"
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoPrt)
'   If bSqlRows Then
'      With RdoPrt
'         Do Until .EOF
'            If "" & Trim(!PartRef) <> Compress(lblAssy) Then _
'               AddComboStr cmbPrt.hwnd, "" & Trim(!PartNum)
'            .MoveNext
'         Loop
'         ClearResultSet RdoPrt
'      End With
'   End If
'   If cmbPrt.ListCount > 0 And BompBMe01a.lstNodes.ListCount > 0 Then
'      For iParts = 0 To cmbPrt.ListCount - 1
'         For iNodes = 0 To BompBMe01a.lstNodes.ListCount - 1
'            If cmbPrt.List(iParts) = BompBMe01a.lstNodes.List(iNodes) Then
'               cmbPrt.RemoveItem iParts
'            End If
'         Next
'      Next
'   End If
'   Set RdoPrt = Nothing
'   bChanged = 0
'   Exit Sub
'
'DiaErr1:
'   BompBMe01a.cmdQuit.Enabled = True
'   BompBMe01a.cmdAdd.Enabled = True
'   BompBMe01a.cmdEdit.Enabled = True
'   BompBMe01a.cmdCut.Enabled = True
'   BompBMe01a.cmdCut.Enabled = True
'   BompBMe01a.cmdCopy.Enabled = True
'   BompBMe01a.cmdDelete.Enabled = True
'   sProcName = "fillpartco"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'
'End Sub
'
Private Sub AddThisPart()
   Dim RdoAdd As ADODB.Recordset
   MouseCursor 13
   cmdAdd.Enabled = False
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT * FROM BmplTable WHERE BMASSYPART='" _
          & Compress(lblAssy) & "' AND BMREV='" & cmbRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAdd, ES_DYNAMIC)
   With RdoAdd
      
      .AddNew
      !BMASSYPART = Compress(lblAssy)
      !BMPARTREF = Compress(cmbPrt)
      !BMPARTNUM = cmbPrt
      !BMREV = Trim(lblRev)
      !BMPARTREV = Trim(cmbRev)
      !BMQTYREQD = Val(txtQty)
      !BMUNITS = txtBum
      !BMCONVERSION = Val(txtCvt)
      !BMSEQUENCE = Val(txtSeq)
      !BMADDER = Val(txtAdr)
      !BMSETUP = Val(txtSup)
      !BMPHANTOM = str$(optPhn.Value)
      '!BMREFERENCE = txtRef
      !BMCOMT = Trim(txtCmt)
      !BMESTLABOR = Val(txtLab)
      !BMESTLABOROH = Val(txtLabOh)
      !BMESTMATERIAL = Val(txtMat)
      !BMESTMATERIALBRD = Val(txtMatbr)
      !BMPICKAT = Val(IIf(txtPickAt = "", 1, txtPickAt))
      .Update
      sSql = "UPDATE BmhdTable SET BMHREVDATE='" _
             & Format(ES_SYSDATE, "mm/dd/yy") & "' WHERE " _
             & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
             & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
   End With
   If clsADOCon.ADOErrNum = 0 Then
      lstAssy.AddItem cmbPrt
      cmbPrt = ""
      txtCmt = ""
      txtQty = "0.000"
      txtCvt = "0.000"
      txtAdr = "0.000"
      txtSup = "0.000"
      txtLab = "0.000"
      txtLabOh = "0.000"
      txtMat = "0.000"
      txtMatbr = "0.000"
      txtBum = "EA"
      txtPickAt = "1"
      
      SysMsg "The Item Was Successfully Added", True
      BompBMe01a.optRefresh = vbChecked
   Else
      MsgBox Trim(Err.Descripton) & vbCrLf _
                  & "Couldn't Add The Item.", _
                  vbExclamation, Caption
   End If
   Set RdoAdd = Nothing
   Unload Me
   
End Sub

Private Sub GetThisPart()
   Dim RdoPrt As ADODB.Recordset
   Dim Units As String
   On Error Resume Next
   sSql = "SELECT PARTREF,PAUNITS,PALEVEL,PATOOL,PASTDCOST FROM PartTable " _
          & "WHERE (PATOOL=0 AND PARTREF='" & Compress(cmbPrt) & "') "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         If Not IsNull(.Fields(1)) Then _
                       Units = "" & Trim(!PAUNITS) Else _
                       Units = "EA"
         lblTyp = "" & Trim(!PALEVEL)
         txtMat = Format(.Fields(3), ES_QuantityDataFormat)
         ClearResultSet RdoPrt
      End With
   Else
      Units = "EA"
   End If
   If txtBum = "" Or txtBum = "EA" Then txtBum = Units
   Set RdoPrt = Nothing
   
End Sub

Private Sub lblAssy_Change()
   GetIndexHeader
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 12) = "*** Part Num" Then _
           lblDsc.ForeColor = ES_RED Else _
           lblDsc.ForeColor = Es_TextForeColor
   
End Sub

Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 9)
   txtAdr = Format(Abs(Val(txtAdr)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtBum_Change()
   bChanged = 1
   
End Sub

Private Sub txtBum_LostFocus()
   txtBum = CheckLen(txtBum, 2)
   If txtBum = "" Then txtBum = "EA"
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
End Sub


Private Sub txtCvt_LostFocus()
   txtCvt = CheckLen(txtCvt, 9)
   txtCvt = Format(Abs(Val(txtCvt)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 9)
   txtLab = Format(Abs(Val(txtLab)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtLabOh_LostFocus()
   txtLabOh = CheckLen(txtLabOh, 9)
   txtLabOh = Format(Abs(Val(txtLabOh)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 9)
   txtMat = Format(Abs(Val(txtMat)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtMatbr_LostFocus()
   txtMatbr = CheckLen(txtMatbr, 9)
   txtMatbr = Format(Abs(Val(txtMatbr)), ES_QuantityDataFormat)
   
End Sub



Private Sub txtQty_Change()
   bChanged = 1
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If lblDsc.ForeColor <> ES_RED Then cmdAdd.Enabled = True
   
End Sub


Private Sub txtSeq_Change()
   bChanged = 1
   
End Sub

Private Sub txtSeq_LostFocus()
   txtSeq = CheckLen(txtSeq, 3)
   txtSeq = Format$(Abs(Val(txtSeq)), "##0")
   
End Sub


Private Function PartOk(ByVal sPartNum As String) As Byte
    Dim iNodes As Integer
    Dim rdoPrtTool As ADODB.Recordset
    
    PartOk = 1
    If Compress(sPartNum) = Compress(lblAssy) Then
        PartOk = 0
        Exit Function
    End If
    For iNodes = 0 To BompBMe01a.lstNodes.ListCount - 1
        If cmbPrt = BompBMe01a.lstNodes.list(iNodes) Then
            PartOk = 0
            Exit Function
        End If
    Next iNodes
    
    sSql = "SELECT PARTREF FROM PartTable WHERE PartREF='" & Compress(sPartNum) & "' AND PATOOL=0"
    If clsADOCon.GetDataSet(sSql, rdoPrtTool, ES_FORWARD) <> 1 Then PartOk = 0
    ClearResultSet rdoPrtTool
    Set rdoPrtTool = Nothing
    
End Function

Function SetPartSearchOption(bPartSearch As Boolean)
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFindPart.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFindPart.Visible = False
   End If
End Function

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = UCase(CheckLen(txtPrt, 30))
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"  '?
   cmbPrt = txtPrt
   
   'get UOM from part
   Dim rs As ADODB.Recordset
   sSql = "select PAUNITS from PartTable where PARTREF = '" & Compress(txtPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   If bSqlRows Then
      txtBum = rs!PAUNITS
   Else
      txtBum = ""
   End If
   Set rs = Nothing
   
End Sub

