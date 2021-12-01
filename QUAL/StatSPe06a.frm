VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPe06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Family ID's"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPe06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6120
      TabIndex        =   2
      ToolTipText     =   "Update Current Part Number And Apply Changes"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbFam 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select From List"
      Top             =   1920
      Width           =   1875
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter Part Number Or Select From List"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2985
      FormDesignWidth =   7065
   End
   Begin VB.Label lblFam 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblCde 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   11
      ToolTipText     =   "Product Code"
      Top             =   1440
      Width           =   870
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prod Code"
      Height          =   285
      Index           =   4
      Left            =   4800
      TabIndex        =   10
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      ToolTipText     =   "Part Level"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   8
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Family ID"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1470
      Width           =   1395
   End
End
Attribute VB_Name = "StatSPe06a"
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
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim bOnLoad As Byte
Dim bGoodPart As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbFam_Click()
   GetFamilyId (cmbFam)
   
End Sub


Private Sub cmbFam_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbFam = CheckLen(cmbFam, 15)
   If Trim(cmbFam) = "" Then
      cmbFam = "NONE"
      b = 1
   Else
      For iList = 0 To cmbFam.ListCount - 1
         If cmbFam = cmbFam.List(iList) Then b = 1
      Next
   End If
   If b = 0 Then
      Beep
      MsgBox "No Such Family ID Has Been Recorded.", _
         vbExclamation, Caption
   End If
   GetFamilyId (cmbFam)
   
End Sub


Private Sub cmbPrt_Click()
   bGoodPart = GetCurrentPart()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodPart = GetCurrentPart()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6306
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sFamily As String
   
   If lblDsc.ForeColor = ES_RED Then
      MsgBox "Function Requires A Valid Part Number.", _
         vbExclamation, Caption
      Exit Sub
   End If
   If Trim(cmbFam) = "" Or Trim(cmbFam) = "NONE" Then
      sMsg = "You Have Chosen To Set This Part With" & vbCr _
             & "No Family ID. Continue To Update?"
      sFamily = ""
   Else
      sMsg = "You Have Chosen To Update This Part With" & vbCr _
             & cmbFam & ". Continue To Update?"
      sFamily = Compress(cmbFam)
   End If
   
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE PartTable SET PAFAMILY='" & sFamily _
             & "' WHERE PARTREF='" & Compress(cmbPrt) & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         MsgBox "Part Family Successfully Updated.", _
            vbInformation, Caption
      Else
         MsgBox "Part Family Was Not Successfully Updated.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT TOP 1 PARTREF,PARTNUM,PADESC,PALEVEL," _
          & "PAPRODCODE,PAFAMILY FROM PartTable " _
          & "WHERE PARTREF= ? "

   
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar
   
   AdoQry.Parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set StatSPe06a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PALEVEL FROM PartTable WHERE " _
          & "(PALEVEL<>5 AND PALEVEL<>6 AND PATOOL=0 AND PAINACTIVE = 0 And PAOBSOLETE = 0) ORDER BY PARTREF"
   LoadComboBox cmbPrt
   
   AddComboStr cmbFam.hwnd, "NONE"
   cmbFam = "NONE"
   sSql = "Qry_FillSPFamily"
   LoadComboBox cmbFam
   
   If cmbFam.ListCount = 1 Then _
                         MsgBox "There Are No Family ID's Available.", _
                         vbInformation, Caption
   If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCurrentPart() As Byte
   Dim AdoPrt As ADODB.Recordset
   Dim sFamily As String
   
   On Error GoTo DiaErr1
   'RdoQry(0) = Compress(cmbPrt)
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(AdoPrt, AdoQry, ES_KEYSET)
   If bSqlRows Then
      With AdoPrt
         cmbPrt = "" & Trim(.Fields(1))
         lblDsc = "" & Trim(.Fields(2))
         lblTyp = "" & Trim(.Fields(3))
         lblCde = "" & Trim(.Fields(4))
         GetFamilyId (Trim(.Fields(5)))
         ClearResultSet AdoPrt
         GetCurrentPart = 1
      End With
   Else
      lblTyp = ""
      lblCde = ""
      lblDsc = "*** No Valid Current Part ***"
      GetCurrentPart = 0
   End If
   Set AdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcurrentp"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** No V" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub


Private Sub GetFamilyId(sFamily As String)
   Dim AdoFam As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT FAMID,FAMDESC FROM RjfmTable " _
          & "WHERE FAMREF='" & Compress(sFamily) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoFam, ES_FORWARD)
   If bSqlRows Then
      With AdoFam
         cmbFam = "" & Trim(.Fields(0))
         lblFam = "" & Trim(.Fields(1))
         ClearResultSet AdoFam
      End With
   Else
      cmbFam = "NONE"
      lblFam = "No Family Selected"
   End If
   Set AdoFam = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getfamilyid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
