VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Tax Codes To Parts"
   ClientHeight    =   2415
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2415
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
      Begin VB.OptionButton optWho 
         Caption         =   "Wholesale"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optRet 
         Caption         =   "Retail"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.ComboBox cmbCode 
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Tag             =   "8"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbSte 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Tag             =   "8"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   315
      Left            =   4440
      Picture         =   "diaARe11a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part"
      Top             =   360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5040
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
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
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARe11a.frx":0342
      PictureDn       =   "diaARe11a.frx":0488
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   1680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2415
      FormDesignWidth =   5970
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Code"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   720
      Width           =   3000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1185
   End
End
Attribute VB_Name = "diaARe11a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
' diaARe11a - Assign B&O tax codes to parts.
'
' Created: (JH)
'
' Revisions:
'   10/16/02 (nth) Intergrated into esifina.
'
'************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodPart As Byte
Dim lForeColor As Long

Dim rdoPart As ADODB.Recordset
Dim AdoQry1 As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Private Sub cmbCode_LostFocus()
   If bGoodPart Then
      With rdoPart
         On Error Resume Next
         If optRet Then
            !PABORTAX = Compress(cmbSte) & Compress(cmbCode)
         Else
            !PABOWTAX = Compress(cmbSte) & Compress(cmbCode)
         End If
         .Update
      End With
      If Err > 0 Then ValidateEdit Me
   End If
End Sub

Private Sub cmbPrt_Click()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) Then bGoodPart = GetPart()
End Sub


Private Sub cmbSte_Click()
   If Not bCancel Then FillCodes
End Sub

Private Sub cmbSte_LostFocus()
   If Not bCancel Then FillCodes
End Sub

Private Sub cmdCan_Click()
   Unload Me
   bCancel = True
End Sub

Private Sub cmdFnd_Click()
   optVew.Value = vbChecked
   ViewParts.Show
End Sub


Private Sub Form_Activate()
   
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      lForeColor = Me.ForeColor
      FillStates
      FillCodes
      FillPartCombo cmbPrt
  '    cmbPrt = cUR.CurrentPart
      bGoodPart = GetPart()
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   
   ' B & O wholesale and retail for part
   sSql = "SELECT PARTREF, PARTNUM, PADESC,PABORTAX,PABOWTAX," _
          & "TxcdTable.TAXCODE AS RetCode, TxcdTable.TAXSTATE AS RetSt," _
          & "TxcdTable_1.TAXCODE AS WhoCode, TxcdTable_1.TAXSTATE AS WhoSt " _
          & "FROM PartTable LEFT OUTER JOIN " _
          & "TxcdTable ON PartTable.PABORTAX = TxcdTable.TAXREF LEFT OUTER JOIN " _
          & "TxcdTable TxcdTable_1 ON PartTable.PABOWTAX = TxcdTable_1.TAXREF " _
          & "Where (PARTREF = ?)"
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   AdoQry1.parameters.Append AdoParameter1
   
   Manage 0
   
   bOnLoad = True
   bCancel = False
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodPart = 1 Then
      cUR.CurrentPart = Trim(cmbPrt)
   End If
   Set rdoPart = Nothing
   Set AdoParameter1 = Nothing
   Set AdoQry1 = Nothing
   FormUnload
   Set diaARe11a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub optRet_Click()
   If bGoodPart Then
      With rdoPart
         cmbSte = "" & Trim(!RetSt)
         cmbCode = "" & Trim(!RetCode)
      End With
   End If
End Sub

Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then
      ' Part search is closing refresh form
      'cmbPrt_LostFocus
   End If
End Sub

Private Sub FillStates()
   Dim rdoSt As ADODB.Recordset
   On Error GoTo DiaErr1
   
   cmbSte.Clear
   sSql = "SELECT DISTINCT TAXSTATE FROM TxcdTable WHERE TAXTYPE = 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSt)
   If bSqlRows Then
      With rdoSt
         While Not .EOF
            cmbSte.AddItem "" & Trim(!taxState)
            .MoveNext
         Wend
         .Cancel
      End With
      cmbSte.ListIndex = 0
   End If
   Set rdoSt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillStates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillCodes()
   Dim rdoCode As ADODB.Recordset
   On Error GoTo DiaErr1
   
   cmbCode.Clear
   sSql = "SELECT TAXCODE FROM TxcdTable WHERE (TAXTYPE = 0)"
   If Trim(cmbSte) <> "" Then
      sSql = sSql & " AND (TAXSTATE = '" & Trim(cmbSte) & "')"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCode)
   
   If bSqlRows Then
      With rdoCode
         While Not .EOF
            cmbCode.AddItem "" & Trim(!taxCode)
            .MoveNext
         Wend
      End With
   End If
   Set rdoCode = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillCodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Function GetPart() As Byte
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   
   AdoQry1.parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(rdoPart, AdoQry1, ES_KEYSET)
   
   If bSqlRows Then
      
      With rdoPart
         ' Fill part descriptions ect.
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc.ForeColor = Me.ForeColor
         lblDsc = "" & Trim(!PADESC)
         If optRet Then
            cmbSte = "" & Trim(!RetSt)
            cmbCode = "" & Trim(!RetCode)
         Else
            cmbSte = "" & Trim(!WhoSt)
            cmbCode = "" & Trim(!WhoCode)
         End If
      End With
      GetPart = 1
   Else
      GetPart = 0
      lblDsc.ForeColor = ES_RED
      lblDsc = "*** No Current Part ***"
   End If
   
   Manage GetPart
   
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub Manage(bOn As Byte)
   ' 1     = Enable all controls
   ' not 1 = Disable all controls
   
   If bOn = 1 Then
      cmbSte.enabled = True
      cmbCode.enabled = True
      
   Else
      cmbSte.enabled = False
      cmbCode.enabled = False
      
   End If
End Sub


Private Sub optWho_Click()
   If bGoodPart Then
      With rdoPart
         cmbSte = "" & Trim(!WhoSt)
         cmbCode = "" & Trim(!WhoCode)
      End With
   End If
End Sub
