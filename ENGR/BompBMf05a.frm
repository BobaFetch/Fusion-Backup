VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BompBMf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Global Bill of Material Part Change"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCanSel 
      Caption         =   "&Clear Selection"
      Height          =   315
      Left            =   1560
      TabIndex        =   14
      ToolTipText     =   "Delete This Revision"
      Top             =   2400
      Width           =   1275
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Select &All"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Delete This Revision"
      Top             =   2400
      Width           =   1275
   End
   Begin VB.ComboBox cmbNPn 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Select Part Number"
      Top             =   1320
      Width           =   3405
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   315
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Delete This Revision"
      Top             =   2760
      Width           =   915
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select All Parts"
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      ToolTipText     =   "Delete This Revision"
      Top             =   1680
      Width           =   1275
   End
   Begin VB.ComboBox cmbRPn 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Select Part Number"
      Top             =   480
      Width           =   3405
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   7200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6900
      FormDesignWidth =   8100
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click The Row To Select A Partnumber to Re-Schedule MO"
      Top             =   2760
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      FixedCols       =   0
      BackColorSel    =   -2147483640
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
   Begin VB.Image Chkno 
      Height          =   180
      Left            =   7080
      Picture         =   "BompBMf05a.frx":07AE
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Chkyes 
      Height          =   180
      Left            =   7080
      Picture         =   "BompBMf05a.frx":0805
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Part Number"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblDesc1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number to Replace"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "BompBMf05a"
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
'Dim RdoPrt As rdoQuery
Dim AdoCmdObj As ADODB.Command
'Dim RdoBmh As rdoResultset
Dim RdoBmh As ADODB.Recordset

Dim bGoodPart As Byte
Dim bGoodList As Byte
Dim bOnLoad As Byte

Dim sPartNumber As String
Dim sPartBomrev As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   
End Sub

Private Sub cmbNPn_Click()
   GetCurrentPart cmbNPn, lblDesc1
End Sub

Private Sub cmbRPn_Click()
   GetCurrentPart cmbRPn, lblDsc
End Sub

Private Sub cmbRPn_LostFocus()
   cmbRPn = CheckLen(cmbRPn, 30)
   
   GetCurrentPart cmbRPn, lblDsc
   ' get the Partlist for the replace partnumner
   Dim iPartType As Integer
   iPartType = GetPartType(Compress(cmbRPn))
   
   If (iPartType <> 0) Then
   
      ClearGrid
      lblType = CStr(iPartType)
      cmbNPn.Clear
      FillPartListByType cmbNPn, iPartType
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub ClearGrid()
      
   grd.Clear
   grd.Rows = 2
   With grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "Assembly Part Number"
      .Col = 2
      .Text = "Req Qty"
      .ColWidth(0) = 700
      .ColWidth(1) = 4000
      .ColWidth(2) = 1200
      
   End With
      
End Sub

Private Sub cmdCanSel_Click()
   Dim iList As Integer
   For iList = 1 To grd.Rows - 1
       grd.Col = 0
       grd.Row = iList
       ' Only if the part is checked
       If grd.CellPicture = Chkyes.Picture Then
           Set grd.CellPicture = Chkno.Picture
       End If
   Next
End Sub

Private Sub cmdSelAll_Click()
   Dim iList As Integer
   For iList = 1 To grd.Rows - 1
       grd.Col = 0
       grd.Row = iList
       ' Only if the part is checked
       If grd.CellPicture = Chkno.Picture Then
           Set grd.CellPicture = Chkyes.Picture
       End If
   Next
End Sub

Private Sub cmdUpdate_Click()
   Dim bAssigned As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sNewRevision As String
   Dim iList As Integer
   Dim strAssmPartNum As String
   Dim strBomQty As String
   
   Dim strNewPartRef As String
   Dim strRepPart As String
   
   strNewPartRef = cmbNPn
   strRepPart = cmbRPn
   
   clsADOCon.BeginTrans
   
   ' Go throught all the record int he grid and re-schedule MO
   For iList = 1 To grd.Rows - 1
      grd.Col = 0
      grd.Row = iList
      ' Only if the part is checked
      If grd.CellPicture = Chkyes.Picture Then
         grd.Col = 1
         strAssmPartNum = grd.Text
        
         sSql = "UPDATE BmplTable SET BMPARTREF = '" & Compress(strNewPartRef) & "', " _
               & " BMPARTNUM = '" & strNewPartRef & "' WHERE " _
                 & " BMASSYPART = '" & strAssmPartNum & "' AND BMPARTREF = '" & Compress(strRepPart) & "'"
      
'         RdoCon.Execute sSql, rdExecDirect
        clsADOCon.ExecuteSQL sSql 'rdExecDirect
     
      End If
    Next
            
   If Err = 0 Then
   
      clsADOCon.CommitTrans
      SysMsg "BOM Sub Part has been replaced.", True, Me
   Else
      clsADOCon.RollbackTrans
      SysMsg "BOM Sub Part couldn't be replaced.", True, Me
   End If

   MouseCursor 0
   Exit Sub
   
BdeleDn1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume BdeleDn2
BdeleDn2:
   On Error Resume Next
   clsADOCon.RollbackTrans
   MouseCursor 0
   MsgBox Err.Description & vbCrLf _
      & "Could Not Delete The Parts List.", vbExclamation, Caption
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3251
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSel_Click()

'   Dim RdoGrd As rdoResultset
   Dim RdoGrd As ADODB.Recordset

   
   On Error Resume Next
   
   Dim strPartReplace As String
   Dim strPartNew As String
   
   grd.Rows = 1
   On Error GoTo DiaErr1
   
   ' IF the begin date and end date are ALL
   ' Get the max and min dates from MRPL Table
   
   strPartReplace = Trim(Compress(cmbRPn))
   strPartNew = Trim(Compress(cmbNPn))
    
   If (strPartReplace <> "" And strPartNew <> "") Then
    
      sSql = "select BMASSYPART, BMPARTREF,  BMQTYREQD from BmplTable" & _
               " WHERE BMPARTREF = '" & strPartReplace & "'"
      
     ' bSqlRows = GetDataSet(RdoGrd, ES_FORWARD)
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoGrd, ES_FORWARD)
     
      If bSqlRows Then
         With RdoGrd
            Do Until .EOF
               grd.Rows = grd.Rows + 1
               grd.Row = grd.Rows - 1
               grd.Col = 0
               Set grd.CellPicture = Chkno.Picture
               grd.Col = 1
               grd.Text = "" & Trim(!BMASSYPART)
               grd.Col = 2
               grd.Text = "" & Trim(!BMQTYREQD)
               
               .MoveNext
            Loop
            ClearResultSet RdoGrd
         End With
      Else
         MsgBox "There Are No Parts for this criteria.", _
            vbInformation, Caption
      End If
      Set RdoGrd = Nothing
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "cmdSel_Click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillPartsBelow4 cmbRPn
      If cmbRPn.ListCount > 0 Then cmbRPn = cmbRPn.List(0)
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub
Private Function GetPartType(strPartNum As String) As Integer
   'Dim RdoType As rdoResultset
   Dim RdoType As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT PALEVEL FROM PartTable WHERE PARTREF = '" & Compress(strPartNum) & "'"
  ' bSqlRows = GetDataSet(RdoType, ES_FORWARD)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoType, ES_FORWARD)
     
    If bSqlRows Then
      GetPartType = RdoType!PALEVEL
      ClearResultSet RdoType
   Else
      GetPartType = -1
   End If
   Set RdoType = Nothing
   Exit Function
   
modErr1:
   sProcName = "GetPartType"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function
Private Sub FillPartListByType(Cntrl As Control, Ptypenum As Integer)
   'Dim RdoFp4 As rdoResultset
   
   Dim RdoFp4 As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT PARTNUM FROM PartTable WHERE PALEVEL = " & Ptypenum
'   bSqlRows = GetDataSet(RdoFp4, ES_FORWARD)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFp4, ES_FORWARD)

   If bSqlRows Then
      With RdoFp4
         Do Until .EOF
            AddComboStr Cntrl.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoFp4
      End With
   End If
   If Cntrl.ListCount > 0 Then Cntrl = Cntrl.List(0)
   Set RdoFp4 = Nothing
   Exit Sub
   
modErr1:
   sProcName = "FillPartListByType"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub


Private Sub Grd_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      grd.Col = 0
      If grd.CellPicture = Chkyes.Picture Then
         Set grd.CellPicture = Chkno.Picture
      Else
         Set grd.CellPicture = Chkyes.Picture
      End If
   End If
   
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   grd.Col = 0
   If grd.CellPicture = Chkyes.Picture Then
      Set grd.CellPicture = Chkno.Picture
   Else
      Set grd.CellPicture = Chkyes.Picture
   End If
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   
   With grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Sel"
      .Col = 1
      .Text = "Assembly Part Number"
      .Col = 2
      .Text = "Req Qty"
      .ColWidth(0) = 700
      .ColWidth(1) = 4000
      .ColWidth(2) = 1200

   End With

   MouseCursor 0
   
   
   bOnLoad = 1
   
End Sub

'
Private Sub Form_Resize()
   Refresh

End Sub


Private Sub Form_Unload(Cancel As Integer)
   SaveCurrentSelections
   On Error Resume Next
   FormUnload
   Set BompBMf05a = Nothing
   
End Sub

