VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Packing Slip Shipped Flag"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "Cancel Packing Slip Shipped Flag"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Printed Packslips Not Printed"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   480
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2145
      FormDesignWidth =   6225
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Function Is For Allowing Prepacked Goods"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1395
      Width           =   3495
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3885
      TabIndex        =   2
      Top             =   1080
      Width           =   1035
   End
End
Attribute VB_Name = "PackPSf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/5/04 Added SoitTable.ITPSSHIPPED
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim bOnLoad As Byte
Dim bGoodPs As Byte

Dim iTotalItems As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbPsl_Click()
   bGoodPs = GetPackslip()
   
End Sub


Private Sub cmbPsl_LostFocus()
   cmbPsl = CheckLen(cmbPsl, 8)
   ' Not need to prepend "PS"
   'If Val(cmbPsl) > 0 Then cmbPsl = Format(cmbPsl, "00000000")
   bGoodPs = GetPackslip()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2252
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdItm_Click()
   Dim b As Byte
   If bGoodPs Then
      CancelShipped
   Else
      MsgBox "Requires A Valid Packing Slip.", _
         vbExclamation, Caption
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillPackSlips
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT PSNUMBER,PSCUST,PSPRINTED FROM PshdTable WHERE PSNUMBER= ? " _
          & "AND (PSSHIPPED=1 AND PSINVOICE=0)"
 '  Set rdoQry = RdoCon.CreateQuery("", sSql)
 '  rdoQry.MaxRows = 1
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adChar
   prmObj.Size = 8
   
   cmdObj.Parameters.Append prmObj
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set cmdObj = Nothing
   Set PackPSf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Function GetPackslip() As Byte
   Dim RdoCsf As ADODB.Recordset
   On Error GoTo DiaErr1
   GetPackslip = False
'   rdoQry.RowsetSize = 1
'   rdoQry(0) = Compress(cmbPsl)
'   bSqlRows = GetQuerySet(RdoCsf, rdoQry, ES_KEYSET)
   
   cmdObj.Parameters(0).Value = Compress(cmbPsl)
   bSqlRows = clsADOCon.GetQuerySet(RdoCsf, cmdObj, ES_FORWARD, True)
   
   If bSqlRows Then
      With RdoCsf
         lblDte = "" & Format(!PSPRINTED, "mm/dd/yyyy")
         If Trim(!PSCUST) <> "" Then FindCustomer Me, Trim(!PSCUST), True
      End With
      ClearResultSet RdoCsf
      GetPackslip = 1
   Else
      lblDte = ""
      lblCst = ""
      lblNme = "*** Invalid Packing Slip ***"
      GetPackslip = 0
   End If
   Set RdoCsf = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpacksl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillPackSlips()
   On Error GoTo DiaErr1
   cmbPsl.Clear
   sSql = "SELECT DISTINCT PSNUMBER FROM " _
          & "PshdTable,PsitTable WHERE " _
          & "(PSSHIPPED=1 AND PSINVOICE=0) "
   LoadComboBox cmbPsl, -1
   If cmbPsl.ListCount > 0 Then
      cmdItm.Enabled = True
      cmbPsl = cmbPsl.List(0)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillpacks"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblNme_Change()
   If Left(lblNme, 5) = "*** I" Then
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = Es_TextForeColor
   End If
   
End Sub


Private Sub CancelShipped()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Do You Really Want To Cancel The " & vbCrLf _
          & "Shipment Of  " & cmbPsl & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE PshdTable SET PSSHIPPEDDATE=NULL, PSSHIPPED=0 " _
             & "WHERE PSNUMBER='" & Trim(cmbPsl) & "'"
       clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "UPDATE SoitTable SET ITPSSHIPPED=0 " _
             & "WHERE ITPSNUMBER='" & Trim(cmbPsl) & "'"
       clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         SysMsg cmbPsl & " Flag Was Changed.", True
         FillPackSlips
      Else
         clsADOCon.RollbackTrans
         MsgBox "Couldn't Change Shipped Flag.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   cmdItm.Enabled = True
   sProcName = "cancelpri"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
