VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Packing Slip"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
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
      ToolTipText     =   "Cancel This Packing Slip Entirely"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Contains Packslips Not Printed"
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
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2205
      FormDesignWidth =   6240
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
Attribute VB_Name = "PackPSf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/30/06 Added ITPSSHIPPED = 0
Option Explicit
'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim RdoCan As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodPs As Byte

Dim iTotalItems As Integer

Dim vItems(800, 13) As Variant
'   0 = PIITNO
'   1 = PITYPE
'   2 = PIQTY
'   3 = PIPART
'   4 = PISONUMBER
'   5 = PISOITEM
'   6 = PISOREV


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbPsl_Click()
   bGoodPs = GetPackslip()
   
End Sub


Private Sub cmbPsl_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   
   cmbPsl = CheckLen(cmbPsl, 8)
   If Val(cmbPsl) > 0 Then cmbPsl = Format(cmbPsl, "00000000")
   If cmbPsl.ListCount > 0 Then
      For iList = 0 To cmbPsl.ListCount - 1
         If cmbPsl.List(iList) = cmbPsl Then bByte = 1
      Next
      If bByte = 0 Then
         Beep
         cmbPsl = cmbPsl.List(0)
      End If
      bGoodPs = GetPackslip()
   'Else
   '   MsgBox "No Packing Slips Qualify.", _
   '      vbInformation, Caption
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2250
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdItm_Click()
   Dim b As Byte
   b = GetItems()
   If bGoodPs Then
      CancelPackSlip
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
   bGoodPs = GetPackslip()
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls

   sSql = "SELECT PSNUMBER,PSCUST,PSDATE FROM PshdTable WHERE PSNUMBER= ? " _
          & "AND (PSTYPE=1 AND PSSHIPPRINT=0 AND PSINVOICE=0 AND PSCANCELED=0)"
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
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set cmdObj = Nothing
   Set RdoCan = Nothing
   
   Set PackPSf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Function GetPackslip() As Byte
   On Error GoTo DiaErr1
   GetPackslip = 0
 '  rdoQry.RowsetSize = 1
 '  rdoQry(0) = Compress(cmbPsl)
 '  bSqlRows = GetQuerySet(RdoCan, rdoQry, ES_KEYSET)
   cmdObj.Parameters(0).Value = Compress(cmbPsl)
   bSqlRows = clsADOCon.GetQuerySet(RdoCan, cmdObj, ES_FORWARD, True)
  
   If bSqlRows Then
      With RdoCan
         lblDte = "" & Format(!PSDATE, "mm/dd/yyyy")
         If Trim(!PSCUST) <> "" Then FindCustomer Me, Trim(!PSCUST), True
      End With
      GetPackslip = 1
   Else
      lblDte = ""
      lblCst = ""
      lblNme = "*** Invalid Or Not Qualifying Packing Slip ***"
      GetPackslip = 0
   End If
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
   sSql = "SELECT DISTINCT PSNUMBER,PSTYPE FROM PshdTable WHERE " _
          & "(PSSHIPPRINT=0 AND PSINVOICE=0 AND PSCANCELED=0) "
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


Private Sub CancelPackSlip()
   Dim bResponse As Byte
   Dim iRow As Integer
   
   Dim cPartCost As Currency
   Dim sMsg As String
   Dim sPackSlip As String
   
   On Error GoTo DiaErr1
   sMsg = "Do You Really Want To Cancel Packing Slip " & cmbPsl & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      MouseCursor 13
      cmdItm.Enabled = False
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      For iRow = 0 To iTotalItems
         sSql = "UPDATE SoitTable SET ITACTUAL=NULL,ITPSNUMBER=''," _
                & "ITPSITEM=0,ITPSCARTON=0,ITPSSHIPNO=0,ITPSSHIPPED=0 WHERE " _
                & "(ITSO=" & Val(vItems(iRow, 4)) & " AND " _
                & "ITNUMBER=" & Val(vItems(iRow, 5)) & " AND " _
                & "ITREV='" & vItems(iRow, 6) & "')"
         clsADOCon.ExecuteSQL sSql ' rdExecDirect
      Next
      sSql = "DELETE FROM PsitTable WHERE PIPACKSLIP='" & Trim(cmbPsl) & "' "
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      
      sSql = "UPDATE PshdTable SET PSCANCELED=1 " _
             & "WHERE PSNUMBER='" & Trim(cmbPsl) & "' "
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      MouseCursor 0
      If clsADOCon.ADOErrNum = 0 Then
         bResponse = MsgBox("Prepared To Cancel. Continue.", _
                     ES_NOQUESTION, Caption)
         If bResponse = vbYes Then
            clsADOCon.CommitTrans
            MouseCursor 0
            SysMsg "Packing Slip " & cmbPsl & " Canceled.", True
            FillPackSlips
            On Error Resume Next
            cmbPsl.SetFocus
         Else
            clsADOCon.RollbackTrans
            CancelTrans
         End If
      Else
         clsADOCon.RollbackTrans
         MsgBox "Couldn't Cancel This Packing Slip.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   cmdItm.Enabled = True
   sProcName = "cancelpac"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetItems() As Byte
   Dim RdoPsl As ADODB.Recordset
   Dim sPackSlip As String
   
   sPackSlip = Compress(cmbPsl)
   Erase vItems
   iTotalItems = -1
   On Error GoTo DiaErr1
   sSql = "SELECT PIITNO,PITYPE,PIQTY,PIPART,PISONUMBER," _
          & "PISOITEM,PISOREV FROM PsitTable WHERE " _
          & "PIPACKSLIP='" & sPackSlip & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsl, ES_FORWARD)
   If bSqlRows Then
      With RdoPsl
         Do Until .EOF
            iTotalItems = iTotalItems + 1
            vItems(iTotalItems, 0) = !PIITNO
            vItems(iTotalItems, 1) = !PITYPE
            vItems(iTotalItems, 2) = !PIQTY
            vItems(iTotalItems, 3) = "" & Trim(!PIPART)
            vItems(iTotalItems, 4) = !PISONUMBER
            vItems(iTotalItems, 5) = !PISOITEM
            vItems(iTotalItems, 6) = "" & Trim(!PISOREV)
            .MoveNext
         Loop
         ClearResultSet RdoPsl
      End With
      GetItems = 1
   Else
      GetItems = False
   End If
   Set RdoPsl = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
