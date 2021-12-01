VERSION 5.00
Begin VB.Form PurcPRf13a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy a Purchase Order"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbNewPO 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopyPO 
      Caption         =   "Copy PO"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Enter PO Number to Copy"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmbCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "New PO:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "PO to Copy:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label lblVnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "PurcPRf13a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bOnLoad As Byte
Dim bGoodPo As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbCan_Click()
   Unload Me
End Sub


Private Sub cmbPon_Click()
   bGoodPo = GetPurchaseOrder(cmbPon)
   
End Sub

Private Sub cmbPon_LostFocus()
   If Len(cmbPon) > 0 Then
      cmbPon = CheckLen(cmbPon, 6)
      cmbPon = Format(Abs(Val(cmbPon)), "000000")
      bGoodPo = GetPurchaseOrder(cmbPon)
   Else
      bGoodPo = False
   End If
   
End Sub

Private Sub cmdCopyPO_Click()
   Dim RdoCpy As ADODB.Recordset
   
   Dim iRow As Integer
   Dim bResponse As Integer
   Dim sComments As String
   Dim sMsg As String
   
   sMsg = "Are You Sure That You Wish To Copy " & vbCrLf _
          & "Purchase Order " & cmbPon & " To " & cmbNewPO & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then Exit Sub
   MouseCursor 13
   iRow = 10
   'prg1.Value = iRow
   'prg1.Visible = True
   'cmdCpy.Enabled = False
   'cmdDis.Enabled = False
   
   On Error Resume Next
   sSql = "DROP TABLE #Poit" 'Just in case
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DROP TABLE #Pohd" 'Just in case
   clsADOCon.ExecuteSQL sSql
   
   On Error GoTo DiaErr1
   sSql = "SELECT PONUMBER FROM PohdTable WHERE PONUMBER=" & Val(cmbPon) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpy, ES_FORWARD)
   If bSqlRows Then
      Err.Clear
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      With RdoCpy
         'Purchase Order Header
         sSql = "SELECT * INTO #Pohd from PohdTable where PONUMBER=" & Val(cmbPon) & " "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE #Pohd SET PONUMBER=" & Val(cmbNewPO) & _
                ",POCREATE='" & Format(Now, "mm/dd/yy") & "', POCANBY='', POCANCELED=NULL, POPRINTED=NULL, POCAN=0, " & _
                " POREVDT='" & Format(Now, "mm/dd/yy") & "' , PORELEASE=0, PODATE='" & Format(Now, "mm/dd/yy") & "', POSHIP='" & Format(Now, "mm/dd/yy") & "' "
         clsADOCon.ExecuteSQL sSql

         sSql = "INSERT INTO PohdTable SELECT * FROM #Pohd "
         clsADOCon.ExecuteSQL sSql
         
         'iRow = 30
         'prg1.Value = iRow
         If clsADOCon.RowsAffected = 0 Then
            MouseCursor 0
            MsgBox "Couldn't Finish Purchase Order Header." & vbCrLf _
               & "Operation Terminated.", vbExclamation, Caption
            On Error Resume Next
            clsADOCon.RollbackTrans
            'cmdCpy.Enabled = True
            'cmdDis.Enabled = True
            'prg1.Visible = False
            Exit Sub
         End If
      End With
      
      sSql = "SELECT * INTO #Poit from PoitTable where PINUMBER=" & Val(cmbPon) & " "
             '& "AND ITCANCELED=0 "
      clsADOCon.ExecuteSQL sSql
      
      sSql = "UPDATE #Poit SET PINUMBER=" & Val(cmbNewPO) & ",PIRELEASE=0, PITYPE=14, PIPDATE='" & Format(Now, "mm/dd/yy") & "' , PIADATE=NULL, " & _
             " PIWIP='', PIONDOC=0, PILOTNUMBER='', PIREJECTED=0, PIWASTE=0, PIINSBY='', PIINSDATE=NULL, PIENTERED='" & Format(Now, "mm/dd/yy") & "', PIODATE=NULL, " & _
             " PIAMT = 0.00, PIAQTY  = 0.00, PIONDOCK = 0, PIRECEIVED=NULL, PIONDOCKINSPECTED=0, PIONDOCKINSPDATE=NULL, PIONDOCKQTYACC=0, PIONDOCKQTYREJ=0, " & _
             " PIONDOCKINSPECTOR='', PIONDOCKCOMMENT='', PIODDELIVERED=0, PIODDELDATE=NULL, PIODDELPSNUMBER='', PIPRESPLITFROM='', PIONDOCKQTYWASTE=0, " & _
             " PIPORIGDATE='" & Format(Now, "mm/dd/yy") & "' "

      Debug.Print sSql
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO PoitTable SELECT * FROM #Poit "
      clsADOCon.ExecuteSQL sSql
      
      'prg1.Value = 100
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         sSql = "UPDATE ComnTable SET CURPONUMBER=" & Val(cmbNewPO) & " "
         clsADOCon.ExecuteSQL sSql
         MsgBox "Purchase Order Copied.", _
            vbInformation, Caption
      Else
         clsADOCon.RollbackTrans
         MsgBox "Could Not Copy The Purchase Order.", _
            vbInformation, Caption
      End If
   End If
   MouseCursor 0
   Set RdoCpy = Nothing
   'sSql = "UPDATE ComnTable SET COLASTPURCHASEORDER = (SELECT MAX(PONUMBER) AS LASTPO FROM PohdTable) WHERE COREF = 1"
   'sSql = "UPDATE ComnTable SET COLASTPURCHASEORDER = (SELECT MAX(PONUMBER) AS LASTPO FROM PohdTable) WHERE COREF = 1"
   sSql = "UPDATE ComnTable SET COLASTPURCHASEORDER = COLASTPURCHASEORDER + 1 WHERE COREF = 1"
   
   clsADOCon.ExecuteSQL sSql
   GetLastPo
   'cmdCpy.Enabled = True
   'cmdDis.Enabled = True
   'prg1.Visible = False
   'GetLastPurchaseOrder
   On Error Resume Next
'   txtOld.SetFocus
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   clsADOCon.RollbackTrans
   sSql = "DROP TABLE #Poit"
   clsADOCon.ExecuteSQL sSql
   
   MouseCursor 0
   'prg1.Visible = False
   MsgBox "Couldn't Copy The Purchase Order.", vbExclamation, Caption

End Sub

Private Sub Form_Activate()
    If bOnLoad Then
       'FillProductCodes
       'FillProductClasses
    End If
    
    bOnLoad = 0
    MouseCursor 0
End Sub


Private Sub Form_Load()
    FormLoad Me
    FillCombo
    'lbAvailableFields.Clear
    'lbExportFields.Clear
    'FillAvailableFields
    bOnLoad = 1
    FormatControls
    GetOptions
    'If lbExportFields.ListCount = 0 And lbAvailableFields.ListCount = 0 Then FillAvailableFields
    'FillAvailableFields
    GetLastPo
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveOptions
End Sub


Private Sub Form_Resize()
    Refresh
End Sub


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormUnload
    Set PurcPRf12a = Nothing
End Sub


Private Sub SaveOptions()
    Dim sOptions As String
    
    sOptions = ""
    'sOptions = sOptions & cbHeaderRow.Value
    
    SaveSetting "Esi2000", "EsiProd", "prf13a", Trim(sOptions)
End Sub

Private Sub GetOptions()
   Dim sOptions As String
    
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "prf13a", sOptions)
   If Len(sOptions) > 0 Then
'        cmbEndDte = Trim(Mid(sOptions, 10, 8))
    Else
'        cbHeaderRow.Value = 1
   End If
    
End Sub


Private Sub FillCombo()
   Dim iList As Integer
   cmbPon.Clear

   On Error GoTo DiaErr1

   'only display PO's with no received or invoiced items
   sSql = "SELECT PONUMBER FROM PohdTable" & vbCrLf _
      & "WHERE NOT EXISTS (SELECT PINUMBER FROM PoitTable" & vbCrLf _
      & "  WHERE PINUMBER = PONUMBER" & vbCrLf _
      & "  AND PITYPE IN (" & IATYPE_PoReceipt & "," & IATYPE_PoInvoiced & "))" & vbCrLf _
      & "AND POCAN = 0" & vbCrLf _
      & "ORDER BY PONUMBER DESC "
      
   LoadComboBox cmbPon, -1
   bGoodPo = GetPurchaseOrder(cmbPon)
   'LoadNumComboBox cmbSrt, "000000"
'   If cmbSrt.ListCount > 0 Then
'      For iList = cmbSrt.ListCount - 1 To 0 Step -1
'         cmbPon.AddItem cmbSrt.List(iList)
'      Next
'      cmbPon = cmbPon.List(0)
'      bGoodPO = GetPurchaseOrder()
'   End If

   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function GetPurchaseOrder(sPONumber As String, Optional DisplayVendor As Boolean = True) As Byte
   Dim RdoPon As ADODB.Recordset
   
   MouseCursor 13
   sSql = "SELECT PONUMBER,POVENDOR,POCAN,VEREF," _
          & "VENICKNAME,VEBNAME FROM PohdTable,VndrTable WHERE " _
          & "(POVENDOR=VEREF AND PONUMBER=" & Val(sPONumber) & " AND " _
          & "POCAN=0)"
   On Error GoTo DiaErr1
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon)
   If bSqlRows Then
      With RdoPon
         If DisplayVendor Then
            cmbPon = Format(!PONUMBER, "000000")
            lblVnd = "" & Trim(!VENICKNAME)
            lblNme = "" & Trim(!VEBNAME)
         End If
         ClearResultSet RdoPon
      End With
      GetPurchaseOrder = True
   Else
      If DisplayVendor Then
        lblVnd = "** Wasn't Found **"
        lblNme = ""
      End If
      GetPurchaseOrder = False
   End If
   On Error Resume Next
   MouseCursor 0
   Set RdoPon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpurchaseor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   GetPurchaseOrder = False
   DoModuleErrors Me
   
End Function

Private Sub GetLastPo()
   Dim RdoCmn As ADODB.Recordset
   Dim lOldPo As Long
       
   On Error GoTo DiaErr1
   lOldPo = 0
   
   sSql = "SELECT COLASTPURCHASEORDER From ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmn, ES_FORWARD)
   
   If bSqlRows Then
      If RdoCmn!COLASTPURCHASEORDER > 0 Then lOldPo = RdoCmn!COLASTPURCHASEORDER
   End If
   
   If lOldPo = 0 Then
      sSql = "SELECT MAX(PONUMBER) AS LASTPO FROM PohdTable "
      Set RdoCmn = clsADOCon.GetRecordSet(sSql)
      'Set RdoCmn = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
      If Not RdoCmn!LASTPO Then
         If RdoCmn!LASTPO > 0 Then lOldPo = RdoCmn!LASTPO
      End If
   End If
   
   cmbNewPO = Format(lOldPo + 1, "000000")
   Set RdoCmn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlastpo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


'Private Sub GetLastPo()
'   Dim RdoCmn As ADODB.Recordset
'
''   Static bNoPos As Byte
''   Dim lOldPo As Long
''   Static sOldLast As String
'    Dim lOldPo As Long
'
'   On Error GoTo DiaErr1
'   lOldPo = 0
'   sSql = "SELECT COLASTPURCHASEORDER From ComnTable WHERE COREF=1"
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmn, ES_FORWARD)
'   If bSqlRows Then
'      If RdoCmn!COLASTPURCHASEORDER > 0 Then lOldPo = RdoCmn!COLASTPURCHASEORDER
'   End If
'
'   If lOldPo = 0 Then
'      sSql = "SELECT MAX(PONUMBER) AS LASTPO FROM PohdTable "
'      Set RdoCmn = clsADOCon.GetRecordSet(sSql)
'      'Set RdoCmn = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
'      If Not RdoCmn!LASTPO Then
'         If RdoCmn!LASTPO > 0 Then lOldPo = RdoCmn!LASTPO
'      'Else
'         '1/24/04
'         'lblLst = "000000"
'         'If bNoPos = 0 Then cmbNewPO = "000001"
'         'bNoPos = 1
'         'tmr1.Enabled = False
'      End If
'   End If
'   'If lOldPO = 0 Then
'      'lblLst = "000000"
'   '   If bNoPos = 0 Then cmbNewPO = "000001"
'   'Else
'      'lblLst = Format(lOldPo, "000000")
'      'If bFillText Then
'         'If sOldLast <> lblLst Then
'   cmbNewPO = Format(lOldPo + 1, "000000")
'      'End If
'   'End If
'   'sOldLast = lblLst
'   Set RdoCmn = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getlastpo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
