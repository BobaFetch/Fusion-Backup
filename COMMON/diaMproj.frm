VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaMproj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Charge Material To A Project"
   ClientHeight    =   4290
   ClientLeft      =   1620
   ClientTop       =   960
   ClientWidth     =   6795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCmt 
      Height          =   765
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "9"
      ToolTipText     =   "Comment (255 Char Max)"
      Top             =   3360
      Width           =   3495
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CheckBox optLot 
      Alignment       =   1  'Right Justify
      Caption         =   "Lot Tracked Part"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      Top             =   3000
      Width           =   1575
   End
   Begin Threed.SSFrame fra1 
      Height          =   30
      Left            =   120
      TabIndex        =   23
      Top             =   1500
      Width           =   6585
      _Version        =   65536
      _ExtentX        =   11606
      _ExtentY        =   53
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Add"
      Height          =   315
      Left            =   5780
      TabIndex        =   6
      ToolTipText     =   "Charge This Item To The Project"
      Top             =   1680
      Width           =   915
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      ToolTipText     =   "Adjustment Quantity"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox cmbPpr 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   " Part Number To Be Charged"
      Top             =   2280
      Width           =   3255
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Select Project Part Number"
      Top             =   720
      Width           =   3545
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5780
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   8
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
      PictureUp       =   "diaMproj.frx":0000
      PictureDn       =   "diaMproj.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   2760
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4290
      FormDesignWidth =   6795
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   1000
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   22
      ToolTipText     =   "Part Type"
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type "
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblCst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uom     "
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
      Left            =   5760
      TabIndex        =   18
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qoh/Chg Qty         "
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
      Left            =   4560
      TabIndex        =   17
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Charged Part Number                                                 "
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
      Left            =   1200
      TabIndex        =   16
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblPsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5760
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4560
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Project"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "diaMproj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/20/04 Fixed InvaTable
'12/23/04 Added Plan Date, reset the Status flag
'1/24/05 Added PKUNITS and exclude Tools
'2/16/05 Fixed cmbRun to properly load Runs
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter


Dim bCanceled As Byte
Dim bFIFO As Byte
Dim bOnLoad As Byte
Dim bGoodMat As Byte
Dim bGoodRuns As Byte

Dim sPartNumber As String
Dim sCreditAcct As String
Dim sDebitAcct As String
Dim sJournalID As String

Dim sLots(50, 2) As String
'0 = Lot Number
'1 = Lot Quantity

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iList As Integer
   Erase sLots
   On Error GoTo DiaErr1
   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartWithLot & "' AND " _
          & "LOTREMAININGQTY>0 AND LOTAVAILABLE=1) "
   If bFIFO = 1 Then
      sSql = sSql & "ORDER BY LOTNUMBER ASC"
   Else
      sSql = sSql & "ORDER BY LOTNUMBER DESC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            iList = iList + 1
            sLots(iList, 0) = "" & Trim(!lotNumber)
            sLots(iList, 1) = Format$(!LOTREMAININGQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         .Cancel
      End With
      GetPartLots = iList
   Else
      GetPartLots = 0
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetPartLots = 0
   
End Function

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = "0.000"
   txtDte = Format(Now, "mm/dd/yy")
   
End Sub

Private Sub cmbPpr_Click()
   bGoodMat = FindMatPart()
   
End Sub

Private Sub cmbPpr_LostFocus()
   cmbPpr = CheckLen(cmbPpr, 30)
   bGoodMat = FindMatPart()
   
End Sub


Private Sub cmbPrt_Click()
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCanceled = 1 Then Exit Sub
   bGoodRuns = GetRuns()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = 1
   
End Sub


'Lots 5/30/02

Private Sub cmdChg_Click()
   PickProject
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs5204"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bFIFO = GetInventoryMethod()
      FillMaterial
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS NOT LIKE 'C%')"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   AdoQry.parameters.Append AdoParameter
   
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set diaMproj = Nothing
   
End Sub



Private Sub FillCombo()
   Dim RdoPrj As ADODB.Recordset
   Dim b As Byte
   Dim sTempPart As String
   
   On Error GoTo DiaErr1
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For This Period.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   sProcName = "fillcombo"
   
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF," _
          & "RUNSTATUS FROM PartTable,RunsTable WHERE PALEVEL=8 " _
          & "AND PARTREF=RUNREF AND (RUNSTATUS<>'CA' OR RUNSTATUS<>'CL') " _
          & "ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrj)
   If bSqlRows Then
      With RdoPrj
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         Do Until .EOF
            If sTempPart <> Trim(!PartNum) Then
               AddComboStr cmbPrt.hWnd, "" & Trim(!PartNum)
               sTempPart = Trim(!PartNum)
            End If
            .MoveNext
         Loop
      End With
      bGoodRuns = GetRuns()
   Else
      cmbPpr.Clear
      lblPsc = ""
      lblQty = ""
      lblUom = ""
      MsgBox "No Project (Part Type 8) Runs Have Been Recorded.", _
         vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoPrj = Nothing
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Function GetRuns() As Byte
   Dim RdoMat As ADODB.Recordset
   Dim iOldLevel As Integer
   iOldLevel = Val(lblTyp)
   cmbRun.Clear
   sPartNumber = GetCurrentPart(cmbPrt, lblDsc)
   lblTyp = iOldLevel
   On Error GoTo DiaErr1
   AdoQry.parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoMat, AdoQry)
   If bSqlRows Then
      With RdoMat
         cmbRun = Format(!Runno, "####0")
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
      End With
      GetRuns = True
   Else
      sPartNumber = ""
      GetRuns = False
   End If
   On Error Resume Next
   Set RdoMat = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillMaterial()
   Dim RdoMat As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL " _
          & "FROM PartTable WHERE (PALEVEL<6 AND PATOOL=0) ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMat)
   If bSqlRows Then
      With RdoMat
         cmbPpr = Trim(!PartNum)
         lblTyp = Format(!PALEVEL, "0")
         Do Until .EOF
            AddComboStr cmbPpr.hWnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
      End With
      bGoodMat = FindMatPart()
   End If
   Set RdoMat = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillmater"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub



Private Function FindMatPart() As Byte
   Dim RdoMat As ADODB.Recordset
   Dim sNewPart
   
   sNewPart = Compress(cmbPpr)
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL,PASTDCOST," _
          & "PAQOH,PALOTTRACK FROM PartTable WHERE PALEVEL<6 " _
          & "AND PARTREF='" & sNewPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMat)
   If bSqlRows Then
      With RdoMat
         cmbPpr = "" & Trim(!PartNum)
         lblPsc = "" & Trim(!PADESC)
         lblUom = "" & !PAUNITS
         lblCst = Format(!PASTDCOST, ES_QuantityDataFormat)
         lblQty = Format(!PAQOH, ES_QuantityDataFormat)
         lblTyp = Format(0 + !PALEVEL, "0")
         optLot.Value = !PALOTTRACK
      End With
      cmdChg.enabled = True
      FindMatPart = True
   Else
      cmdChg.enabled = False
      FindMatPart = False
   End If
   On Error Resume Next
   Set RdoMat = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "findmatpa"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub GetAccounts(sPartNumber As String)
   Dim rdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   
   On Error GoTo DiaErr1
   'Use current Part
   sSql = "Qry_GetExtPartAccounts '" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         bType = Format(!PALEVEL, "0")
         If bType = 6 Or bType = 7 Then
            sDebitAcct = "" & Trim(!PACGSEXPACCT)
            sCreditAcct = "" & Trim(!PAINVEXPACCT)
         Else
            sDebitAcct = "" & Trim(!PACGSMATACCT)
            sCreditAcct = "" & Trim(!PAINVMATACCT)
         End If
         .Cancel
      End With
   Else
      sCreditAcct = ""
      sDebitAcct = ""
      Set rdoAct = Nothing
      Exit Sub
   End If
   Set rdoAct = Nothing
   If sDebitAcct = "" Or sCreditAcct = "" Then
      'None in one or both there, try Product code
      sSql = "Qry_GetPCodeAccounts '" & sPcode & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
      If bSqlRows Then
         With rdoAct
            If bType = 6 Or bType = 7 Then
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCCGSEXPACCT)
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCINVEXPACCT)
            Else
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCCGSMATACCT)
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCINVMATACCT)
            End If
            .Cancel
         End With
      End If
      Set rdoAct = Nothing
      If sDebitAcct = "" Or sCreditAcct = "" Then
         'Still none, we'll check the common
         If bType = 6 Or bType = 7 Then
            sSql = "SELECT COREF,COCGSEXPACCT" & Trim(str(bType)) & "," _
                   & "COINVEXPACCT" & Trim(str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         Else
            sSql = "SELECT COREF,COCGSMATACCT" & Trim(str(bType)) & "," _
                   & "COINVMATACCT" & Trim(str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         End If
         bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
         If bSqlRows Then
            With rdoAct
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(.Fields(0))
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(.Fields(1))
               .Cancel
            End With
         End If
      End If
   End If
   'After this excercise, we'll give up if none are found
   Set rdoAct = Nothing
   Exit Sub
   
DiaErr1:
   'Just bail for now. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
   
End Sub

Private Sub PickProject()
   Dim bResponse As Byte
   Dim bLotsRqd As Byte
   Dim bLotFail As Byte
   Dim a As Integer
   Dim iList As Integer
   Dim iLots As Integer
   Dim iPkRecord As Integer
   Dim lCOUNTER As Long
   Dim lLOTRECORD As Long
   Dim cLotQty As Currency
   Dim cPckQty As Currency
   Dim cItmLot As Currency
   Dim cQuantity As Currency
   Dim cRemPQty As Currency
   
   Dim sDate As String
   Dim sLot As String
   Dim sMsg As String
   Dim sMoRun As String * 9
   Dim sMoPart As String * 31
   Dim sNewPart As String
   
   sDate = Format(ES_SYSDATE, "mm/dd/yy")
   On Error Resume Next
   If Val(txtQty) = 0 Then
      MsgBox "You Have Entered a Zero Quantity.", vbInformation, Caption
      txtQty.SetFocus
      Exit Sub
   Else
      bResponse = CheckLotStatus()
      If bLotsRqd = 1 And optLot.Value = 1 Then
         If Val(txtQty) > Val(lblQty) Then
            MsgBox "This Part Number Is Lot Tracked And There" & vbCr _
               & "Aren't Enough On Hand To Satisfy The Need.", _
               vbInformation, Caption
            Exit Sub
         End If
      End If
      sMsg = "You Have Chosen To Charge " & txtQty & " " & lblUom & vbCr _
             & "Part Number " & cmbPpr & " To The Project." & vbCr _
             & "Do You Wish To Continue?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         MouseCursor 13
         On Error GoTo DiaErr1
         cQuantity = Format(Val(txtQty), ES_QuantityDataFormat)
         sNewPart = Compress(cmbPpr)
         cmdChg.enabled = False
         iList = Len(Trim(str(cmbRun)))
         iList = 5 - iList
         sMoPart = cmbPrt
         sMoRun = "RUN" & Space$(iList) & cmbRun
         sPartNumber = Compress(cmbPrt)
         GetAccounts sNewPart
         On Error Resume Next
         lCOUNTER = GetLastActivity() + 1
         
         clsADOCon.BeginTrans
         clsADOCon.ADOErrNum = 0
         
         iPkRecord = GetNextPickRecord(sPartNumber, Val(cmbRun))
         sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN," _
                & "PKTYPE,PKPDATE,PKADATE,PKPQTY,PKAQTY,PKRECORD,PKUNITS,PKCOMT) " _
                & "VALUES('" & sNewPart & "','" & sPartNumber & "'," _
                & cmbRun & ",10,'" & sDate & "','" & sDate & "'," _
                & cQuantity & "," & cQuantity & "," & iPkRecord & ",'" _
                & lblUom & "','" & txtCmt & "')"
         clsADOCon.ExecuteSQL sSql
         
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPDATE,INPQTY," _
                & "INAQTY,INAMT,INCREDITACCT,INDEBITACCT,INMOPART,INMORUN,INNUMBER," _
                & "INUSER) VALUES(10,'" & sNewPart & "','PICK','" _
                & sMoPart & sMoRun & "','" & txtDte & "',-" & cQuantity & ",-" & cQuantity & "," _
                & lblCst & ",'" & sCreditAcct & "','" & sDebitAcct & "','" _
                & sPartNumber & "'," & Val(cmbRun) & "," & lCOUNTER & ",'" _
                & sInitials & "')"
         clsADOCon.ExecuteSQL sSql
         
         'lots
         iLots = GetPartLots(sNewPart)
         If bLotsRqd = 1 And optLot.Value = vbChecked Then
            'Reqd and Get The lots
            LotSelect.lblPart = Trim(cmbPpr)
            LotSelect.lblRequired = Abs(cQuantity)
            LotSelect.Show vbModal
            If Es_TotalLots > 0 Then
               For a = 1 To UBound(lots)
                  'insert lot transaction here
                  lLOTRECORD = GetNextLotRecord(lots(a).LotSysId)
                  sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                         & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                         & "LOIMOPARTREF,LOIMORUNNO," _
                         & "LOIACTIVITY,LOICOMMENT) " _
                         & "VALUES('" & lots(a).LotSysId & "'," _
                         & lLOTRECORD & ",10,'" & sNewPart & "',-" _
                         & lots(a).LotSelQty & ",'" & sMoPart & "'," & Val(sMoRun) & "," _
                         & lCOUNTER & ",'Project Picked Item')"
                  clsADOCon.ExecuteSQL sSql
                  
                  sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                         & "-" & lots(a).LotSelQty & " WHERE LOTNUMBER='" _
                         & lots(a).LotSysId & "'"
                  clsADOCon.ExecuteSQL sSql
                  cItmLot = cItmLot + lots(a).LotSelQty
               Next
            Else
               bLotFail = 1
            End If
         Else
            For a = 1 To iLots
               cLotQty = Val(sLots(a, 1))
               If cLotQty >= cRemPQty Then
                  cPckQty = cRemPQty
                  cLotQty = cLotQty - cRemPQty
                  cRemPQty = 0
               Else
                  cPckQty = cLotQty
                  cRemPQty = cRemPQty - cLotQty
                  cLotQty = 0
               End If
               cItmLot = cItmLot + cPckQty
               If cItmLot > Val(sLots(a, 1)) Then cItmLot = Val(sLots(a, 1))
               sLot = sLots(a, 0)
               lLOTRECORD = GetNextLotRecord(sLot)
               
               'insert lot transaction here
               sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
                      & "LOITYPE,LOIPARTREF,LOIQUANTITY," _
                      & "LOIMOPARTREF,LOIMORUNNO," _
                      & "LOIACTIVITY,LOICOMMENT) " _
                      & "VALUES('" & sLots(a, 0) & "'," _
                      & lLOTRECORD & ",10,'" & sNewPart & "',-" _
                      & Abs(cItmLot) & ",'" & sMoPart & "'," & Val(sMoRun) & "," _
                      & lCOUNTER & ",'Material To Project')"
               clsADOCon.ExecuteSQL sSql
               
               'Update Lot Header
               sSql = "UPDATE LohdTable SET LOTREMAININGQTY=LOTREMAININGQTY" _
                      & "-" & Abs(cItmLot) & " WHERE LOTNUMBER='" & sLots(a, 0) & "'"
               clsADOCon.ExecuteSQL sSql
               If cRemPQty <= 0 Then Exit For
            Next
         End If
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & Abs(cQuantity) & "," _
                & "PALOTQTYREMAINING=PALOTQTYREMAINING-" & Abs(cItmLot) & " " _
                & "WHERE PARTREF='" & sNewPart & "' "
         clsADOCon.ExecuteSQL sSql
         
         sSql = "UPDATE RunsTable SET RUNSTATUS='PP' WHERE RUNREF='" _
                & Compress(cmbPrt) & "' AND RUNNO=" & Val(cmbRun) & " "
         clsADOCon.ExecuteSQL sSql
         
         MouseCursor 0
         Erase lots()
         If clsADOCon.ADOErrNum = 0 And bLotFail = 0 Then
            clsADOCon.CommitTrans
            AverageCost sNewPart
            
            MsgBox "Material Successfully Charged To The Project.", vbInformation, Caption
            txtQty = ""
            lblQty = ""
            lblUom = ""
            lblCst = ""
            txtCmt = ""
            On Error Resume Next
            cmbRun.SetFocus
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            sMsg = CurrError.Description & vbCr _
                   & "Could Not Complete Project Charge."
            MsgBox sMsg, vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   Exit Sub
   
DiaErr1:
   MouseCursor 0
   CurrError.Description = Err.Description
   On Error Resume Next
   clsADOCon.RollbackTrans
   sMsg = CurrError.Description & vbCr _
          & "Could Not Complete Project Charge."
   MsgBox sMsg, vbExclamation, Caption
   
End Sub