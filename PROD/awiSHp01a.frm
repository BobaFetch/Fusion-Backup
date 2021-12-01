VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form awiShopSHp01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Orders"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "awiSHp01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "awiSHp01a.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optFrom 
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Revision-Select From List"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox optLst 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      ToolTipText     =   "Pick List For This Part (Printed MO's Only) Status PL"
      Top             =   4200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   726
   End
   Begin VB.CheckBox optDoc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "Document List (Printed MO's Only)"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6240
      TabIndex        =   19
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   600
         Picture         =   "awiSHp01a.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "awiSHp01a.frx":0AC2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   12
      Left            =   2760
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optInc 
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   6
      Top             =   2760
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.ComboBox cmbRun 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Select Run Number"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Part Number"
      Top             =   1320
      Width           =   3545
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   3600
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3525
      FormDesignWidth =   7515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick List Information"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PL Rev"
      Height          =   255
      Index           =   18
      Left            =   5160
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   26
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblName 
      Caption         =   "Austin Waterjet Custom Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label lblQty 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick List For This Part"
      Height          =   255
      Index           =   17
      Left            =   480
      TabIndex        =   23
      ToolTipText     =   "Pick List For This Part (Printed MO's Only) Status PL"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblSta 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6840
      TabIndex        =   22
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblTyp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   21
      Top             =   1680
      Width           =   300
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type/Status"
      Height          =   255
      Index           =   15
      Left            =   5160
      TabIndex        =   20
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Information"
      Enabled         =   0   'False
      Height          =   255
      Index           =   14
      Left            =   480
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cover Sheet (Printed MO's Only):"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Document List For This Part"
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   13
      ToolTipText     =   "Document List (Printed MO's Only)"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Allocations"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "awiShopSHp01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/13/02 Added PKRECORD for new index
'9/7/06 Dropped Jet Completely - Added CreateSQL(nn)Allocations
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bGoodPart As Byte
Dim bGoodMo As Byte
Dim bOnLoad As Byte
Dim bTablesCreated As Byte

Dim sBomRev As String
Dim sRunPkstart As String
Dim sPartNumber As String

Dim CustPo(6) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetRevisions()
   cmbRev.Clear
   sSql = "SELECT BMHREV FROM BmhdTable WHERE BMHREF='" _
          & Compress(cmbPrt) & "' ORDER BY BMHREV"
   LoadComboBox cmbRev, -1
   Exit Sub
   
DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   On Error Resume Next
   lblPrinter = GetSetting("Esi2000", "EsiProd", "sh01Printer", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "sh01all", Trim(optInc(5).Value)
   SaveSetting "Esi2000", "EsiProd", "sh01Printer", lblPrinter
   
End Sub




Private Sub cmbPrt_Click()
   bGoodPart = GetRuns()
   If bGoodPart Then GetRevisions
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   bGoodPart = GetRuns()
   If bGoodPart Then GetRevisions
   
End Sub

Private Sub cmbRev_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbRev = CheckLen(cmbRev, 4)
   For iList = 0 To cmbRev.ListCount - 1
      If Trim(cmbRev) = Trim(cmbRev.List(iList)) Then b = 1
   Next
   If b = 0 And cmbRev.ListCount > 0 Then
      Beep
      cmbRev = cmbRev.List(0)
   End If
   
End Sub


Private Sub cmbRun_Click()
   GetThisRun
   
End Sub


Private Sub cmbRun_LostFocus()
   cmbRun = CheckLen(cmbRun, 5)
   If Val(cmbRun) > 32767 Then cmbRun = "32767"
   cmbRun = Format(Abs(Val(cmbRun)), "####0")
   GetThisRun
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      'SelectHelpTopic Me, Caption
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      CreateSQLSOAllocations
      CreateSQLPOAllocations
      CreateSQLPLAllocations
      FillAllRuns cmbPrt
      If optFrom.Value = vbChecked Then
         cmbPrt = ShopSHe02a.cmbPrt
         cmbRun = ShopSHe02a.cmbRun
      End If
      If cmbPrt.ListCount > 0 Then bGoodPart = GetRuns()
      If bGoodPart Then GetRevisions
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bTablesCreated = 0
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PARUN,RUNREF,RUNSTATUS," _
          & "RUNNO FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   AdoQry.Parameters.Append AdoParameter
   
           'Set rdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = 1
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   
   If optFrom.Value = vbChecked Then
      ShopSHe02a.lblStat = lblSta
      ShopSHe02a.Show
   Else
      FormUnload
   End If
   Set awiShopSHp01a = Nothing
   
End Sub




Private Function GetRuns() As Byte
   Dim RdoRns As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbRun.Clear
   MouseCursor 13
   sPartNumber = Compress(cmbPrt)
   AdoQry.Parameters(0).Value = sPartNumber
   
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoQry)
   If bSqlRows Then
      With RdoRns
         If optFrom Then
            cmbRun = ShopSHe02a.cmbRun
         Else
            cmbRun = Format(!Runno, "####0")
         End If
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(!PALEVEL, "#")
         Do Until .EOF
            AddComboStr cmbRun.hwnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      GetRuns = True
      GetThisRun
   Else
      sPartNumber = ""
      GetRuns = False
   End If
   MouseCursor 0
   Set RdoRns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblQty_Click()
   'run qty
   
End Sub



Private Sub optDis_Click()
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbPrt.SetFocus
      Exit Sub
   Else
      PrintSpecial
   End If
   
End Sub

Private Sub optDoc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFrom_Click()
   'dummy to check if from Revise mo
   
End Sub



Private Sub optInc_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optLst_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   Dim b As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbPrt.SetFocus
      Exit Sub
   Else
      On Error Resume Next
      If optLst.Value = vbChecked Then
         If lblSta = "SC" Or lblSta = "RL" Then
            sMsg = "Do You Want To Print The MO Pick " & vbCr _
                   & "List And Move The Run Status To PL?"
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               'Build the pick list and change status
               b = 1
               MouseCursor 13
               BuildPartsList
            Else
               CancelTrans
            End If
         Else
            MouseCursor 13
            b = 1
            BuildPickList
         End If
      End If
      '        If optDoc = vbChecked Then
      '            MouseCursor 13
      '            b = 1
      '            BuildDocumentList
      '        End If
      '       If b = 1 Then PrintCover
      PrintSpecial
   End If
   
End Sub



Private Sub GetThisRun()
   Dim RdoRun As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNSTATUS,RUNPKSTART,RUNQTY FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbPrt) & "' AND " _
          & "RUNNO=" & Val(cmbRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblSta = "" & Trim(!RUNSTATUS)
         If lblSta = "SC" Or lblSta = "RL" Then
            cmbRev.Enabled = True
            'optInc(0).Enabled = False
            'optInc(0).Value = vbUnchecked
         Else
            If lblSta <> "CA" Then
               cmbRev.Enabled = False
               ' optInc(0).Enabled = True
               ' optInc(0).Value = vbChecked
            Else
               cmbRev.Enabled = False
               ' optInc(0).Enabled = False
               ' optInc(0).Value = vbUnchecked
            End If
         End If
         If Not IsNull(!RUNPKSTART) Then
            sRunPkstart = Format(!RUNPKSTART, "mm/dd/yy")
         Else
            sRunPkstart = Format(ES_SYSDATE, "mm/dd/yy")
         End If
         lblQty = Format(!RUNQTY, ES_QuantityDataFormat)
         ClearResultSet RdoRun
      End With
   End If
   Set RdoRun = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




'CJS Revised 9/6/06
'SC...no Pick List yet

Private Sub BuildPartsList()
   Dim RdoLst As ADODB.Recordset
   Dim RdoIns As ADODB.Recordset
   Dim b As Byte
   
   Dim iPkRecord As Integer
   Dim cConversion As Currency
   Dim cQuantity As Currency
   Dim cSetup As Currency
   Dim sComment As String
   Dim sUnits As String
   On Error Resume Next
   
   sSql = "TRUNCATE TABLE AwiPLAllocations"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "SELECT * FROM AwiPLAllocations"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_KEYSET)
   
   sSql = "SELECT DISTINCT BMASSYPART FROM BmplTable " _
          & "WHERE BMASSYPART='" & Compress(cmbPrt) & "' " _
          & "AND BMREV='" & Trim(cmbRev) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      b = 1
      ClearResultSet RdoLst
   Else
      MouseCursor 0
      b = 0
      MsgBox "This Part Does Not Have A Parts List.", _
         vbInformation, Caption
   End If
   If b = 1 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALOCATION," _
             & "BMASSYPART,BMPARTREF,BMREV,BMQTYREQD,BMSETUP,BMADDER," _
             & "BMCONVERSION,BMUNITS,BMCOMT FROM PartTable,BmplTable WHERE (" _
             & "PARTREF=BMPARTREF AND BMASSYPART='" & Compress(cmbPrt) & "' " _
             & "AND BMREV='" & sBomRev & "')"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
      If bSqlRows Then
         b = 0
         With RdoLst
            On Error Resume Next
            RdoIns.AddNew
            RdoIns!PLSRunPart = "" & Compress(cmbPrt)
            
            '                   clsAdoCon.BeginTrans
            Do Until .EOF
               If Not IsNull(!BMSETUP) Then
                  cSetup = !BMSETUP
               Else
                  cSetup = 0
               End If
               cQuantity = Format(!BMQTYREQD + !BMADDER, ES_QuantityDataFormat)
               cConversion = Format(!BMCONVERSION, "#####0.0000")
               If cConversion = 0 Then cConversion = 1
               cQuantity = cQuantity / cConversion
               cQuantity = (cQuantity * Val(lblQty)) + cSetup
               b = b + 1
               sComment = "" & Trim(!BMCOMT)
               If Len(sComment) > 255 Then sComment = Left$(sComment, 255)
               sComment = ReplaceString(sComment)
               sUnits = "" & Trim(!BMUNITS)
               If b < 7 Then
                  Select Case b
                     Case 1
                        RdoIns!PLSPart1 = "" & Trim(!PartNum)
                        RdoIns!PLSDesc1 = "" & Trim(!PADESC)
                        RdoIns!PLSPqty1 = cQuantity
                        RdoIns!PLSUom1 = "" & Trim(!BMUNITS)
                        RdoIns!PLSLoc1 = "" & Trim(!PALOCATION)
                        RdoIns!PLSCom1 = sComment
                     Case 2
                        RdoIns!PLSPart2 = "" & Trim(!PartNum)
                        RdoIns!PLSDesc2 = "" & Trim(!PADESC)
                        RdoIns!PLSPqty2 = cQuantity
                        RdoIns!PLSUom2 = "" & Trim(!BMUNITS)
                        RdoIns!PLSLoc2 = "" & Trim(!PALOCATION)
                        RdoIns!PLSCom2 = sComment
                     Case 3
                        RdoIns!PLSPart3 = "" & Trim(!PartNum)
                        RdoIns!PLSDesc3 = "" & Trim(!PADESC)
                        RdoIns!PLSPqty3 = cQuantity
                        RdoIns!PLSUom3 = "" & Trim(!BMUNITS)
                        RdoIns!PLSLoc3 = "" & Trim(!PALOCATION)
                        RdoIns!PLSCom3 = sComment
                     Case Else
                        RdoIns!PLSPart4 = "" & Trim(!PartNum)
                        RdoIns!PLSDesc4 = "" & Trim(!PADESC)
                        RdoIns!PLSPqty4 = cQuantity
                        RdoIns!PLSUom4 = "" & Trim(!BMUNITS)
                        RdoIns!PLSLoc4 = "" & Trim(!PALOCATION)
                        RdoIns!PLSCom4 = sComment
                        
                        '                                Case 5
                        '                                    DbPls!PLSPart5 = "" & Trim(!PARTNUM)
                        '                                    DbPls!PLSDesc5 = "" & Trim(!PADESC)
                        '                                    DbPls!PLSPqty5 = cQuantity
                        '                                    DbPls!PLSUom5 = "" & Trim(!BMUNITS)
                        '                                    DbPls!PLSLoc5 = "" & Trim(!PALOCATION)
                        '                                    DbPls!PLSCom5 = sComment
                        '                                Case 6
                        '                                    DbPls!PLSPart6 = "" & Trim(!PARTNUM)
                        '                                    DbPls!PLSDesc6 = "" & Trim(!PADESC)
                        '                                    DbPls!PLSPqty6 = cQuantity
                        '                                    DbPls!PLSUom6 = "" & Trim(!BMUNITS)
                        '                                    DbPls!PLSLoc6 = "" & Trim(!PALOCATION)
                        '                                    DbPls!PLSCom6 = sComment
                  End Select
               End If
               .MoveNext
            Loop
            ClearResultSet RdoLst
            RdoIns.Update
            ClearResultSet RdoLst
         End With
      Else
         On Error Resume Next
         RdoIns.AddNew
         RdoIns!PLSRunPart = "" & Compress(cmbPrt)
         RdoIns!PLSPart1 = "*** No Parts List Found ***"
         RdoIns.Update
      End If
   Else
      RdoIns.AddNew
      RdoIns!PLSRunPart = "" & Compress(cmbPrt)
      RdoIns!PLSPart1 = "*** No Parts List Recorded ***"
      RdoIns.Update
   End If
   On Error Resume Next
   Set RdoLst = Nothing
   Set RdoIns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "buildpartslist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


'CJS Revised 9/6/06
'Pick list is active

Private Sub BuildPickList()
   Dim RdoLst As ADODB.Recordset
   Dim RdoIns As ADODB.Recordset
   Dim b As Byte
   
   On Error Resume Next
   sSql = "TRUNCATE TABLE AwiPLAllocations"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "SELECT * FROM AwiPLAllocations"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_KEYSET)
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALOCATION," _
          & "PKPARTREF,PKMOPART,PKPQTY,PKADATE,PKAQTY,PKCOMT,PKUNITS FROM PartTable," _
          & "MopkTable WHERE (PARTREF=PKPARTREF AND PKMOPART='" _
          & Compress(cmbPrt) & "' AND PKMORUN=" & cmbRun & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_STATIC)
   If bSqlRows Then
      With RdoLst
         RdoIns.AddNew
         RdoIns!PLSRunPart = "" & Compress(cmbPrt)
         Do Until .EOF
            b = b + 1
            If b < 5 Then
               Select Case b
                  Case 1
                     RdoIns!PLSPart1 = "" & Trim(!PartNum)
                     RdoIns!PLSDesc1 = "" & Trim(!PADESC)
                     RdoIns!PLSPqty1 = Format(!PKPQTY, ES_QuantityDataFormat)
                     RdoIns!PLSAqty1 = Format(!PKAQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        RdoIns!PLSADate1 = Format$(!PKADATE, "mm/dd/yy")
                     End If
                     RdoIns!PLSUom1 = "" & Trim(!PKUNITS)
                     RdoIns!PLSLoc1 = "" & Trim(!PALOCATION)
                     RdoIns!PLSCom1 = "" & Trim(!PKCOMT)
                     
                  Case 2
                     RdoIns!PLSPart2 = "" & Trim(!PartNum)
                     RdoIns!PLSDesc2 = "" & Trim(!PADESC)
                     RdoIns!PLSPqty2 = Format(!PKPQTY, ES_QuantityDataFormat)
                     RdoIns!PLSAqty2 = Format(!PKAQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        RdoIns!PLSADate2 = Format$(!PKADATE, "mm/dd/yy")
                     End If
                     RdoIns!PLSUom2 = "" & Trim(!PKUNITS)
                     RdoIns!PLSLoc2 = "" & Trim(!PALOCATION)
                     RdoIns!PLSCom2 = "" & Trim(!PKCOMT)
                     
                  Case 3
                     RdoIns!PLSPart3 = "" & Trim(!PartNum)
                     RdoIns!PLSDesc3 = "" & Trim(!PADESC)
                     RdoIns!PLSPqty3 = Format(!PKPQTY, ES_QuantityDataFormat)
                     RdoIns!PLSAqty3 = Format(!PKAQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        RdoIns!PLSADate3 = Format$(!PKADATE, "mm/dd/yy")
                     End If
                     RdoIns!PLSUom3 = "" & Trim(!PKUNITS)
                     RdoIns!PLSLoc3 = "" & Trim(!PALOCATION)
                     RdoIns!PLSCom3 = "" & Trim(!PKCOMT)
                     
                  Case Else
                     RdoIns!PLSPart4 = "" & Trim(!PartNum)
                     RdoIns!PLSDesc4 = "" & Trim(!PADESC)
                     RdoIns!PLSPqty4 = Format(!PKPQTY, ES_QuantityDataFormat)
                     RdoIns!PLSAqty4 = Format(!PKAQTY, ES_QuantityDataFormat)
                     If Not IsNull(!PKADATE) Then
                        RdoIns!PLSADate4 = Format$(!PKADATE, "mm/dd/yy")
                     End If
                     RdoIns!PLSUom4 = "" & Trim(!PKUNITS)
                     RdoIns!PLSLoc4 = "" & Trim(!PALOCATION)
                     RdoIns!PLSCom4 = "" & Trim(!PKCOMT)
                     
                     '                            Case 5
                     '                                DbPls!PLSPart5 = "" & Trim(!PARTNUM)
                     '                                DbPls!PLSDesc5 = "" & Trim(!PADESC)
                     '                                DbPls!PLSPqty5 = Format(!PKPQTY, ES_QuantityDataFormat)
                     '                                If Not IsNull(!PKADATE) Then
                     '                                    DbPls!PLSADate5 = Format(!PKADATE, "mm/dd/yy")
                     '                                End If
                     '                                DbPls!PLSAqty5 = Format(!PKAQTY, ES_QuantityDataFormat)
                     '                                DbPls!PLSUom5 = "" & Trim(!PKUNITS)
                     '                                DbPls!PLSLoc5 = "" & Trim(!PALOCATION)
                     '                                DbPls!PLSCom5 = "" & Trim(!PKCOMT)
                     '                            Case 6
                     '                                DbPls!PLSPart6 = "" & Trim(!PARTNUM)
                     '                                DbPls!PLSDesc6 = "" & Trim(!PADESC)
                     '                                DbPls!PLSPqty6 = Format(!PKPQTY, ES_QuantityDataFormat)
                     '                                If Not IsNull(!PKADATE) Then
                     '                                    DbPls!PLSADate6 = Format(!PKADATE, "mm/dd/yy")
                     '                                End If
                     '                                DbPls!PLSAqty6 = Format(!PKAQTY, ES_QuantityDataFormat)
                     '                                DbPls!PLSUom6 = "" & Trim(!PKUNITS)
                     '                                DbPls!PLSLoc6 = "" & Trim(!PALOCATION)
                     '                                DbPls!PLSCom6 = "" & Trim(!PKCOMT)
               End Select
            End If
            .MoveNext
         Loop
         ClearResultSet RdoLst
         RdoIns.Update
      End With
   Else
      RdoIns.AddNew
      RdoIns!PLSRunPart = "" & Compress(cmbPrt)
      RdoIns!PLSPart1 = "*** No Pick List Recorded ***"
      RdoIns.Update
   End If
   On Error Resume Next
   Set RdoLst = Nothing
   Set RdoIns = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "buildpicklist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


'CJS Revised 9/6/06

Private Sub PrintSpecial()
   Dim b As Byte
   Dim sListType As String
   Dim sRout As String
   Dim sRev As String
   
   MouseCursor 13
   Erase CustPo
   On Error GoTo Psh01
   sProcName = "getmorout"
   b = GetMoRouting(sRout, sRev)
   sProcName = "buildall"
   cmdCan.Enabled = False
   BuildPoTable
   BuildAllocations
   If lblSta = "RL" Or lblSta = "SC" Then
      sListType = "Parts List (SC Or RL)"
      BuildPartsList
   Else
      sListType = "Pick List"
      BuildPickList
   End If
   DoEvents
   Sleep 500
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   sCustomReport = GetCustomReport("awish01.rpt")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Routing"
   aFormulaName.Add "RoutingRev"
   aFormulaName.Add "SoPon1"
   
   aFormulaName.Add "SoPon2"
   aFormulaName.Add "SoPon1"
   aFormulaName.Add "SoPon3"
   aFormulaName.Add "SoPon4"
   aFormulaName.Add "SoPon5"
   aFormulaName.Add "ListType"
   aFormulaName.Add "ShowInc"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'" & CStr(sRout) & "'")
   aFormulaValue.Add CStr("'" & CustPo(1) & "'")
   aFormulaValue.Add CStr("'" & CustPo(2) & "'")
   aFormulaValue.Add CStr("'" & CustPo(3) & "'")
   aFormulaValue.Add CStr("'" & CustPo(4) & "'")
   aFormulaValue.Add CStr("'" & CustPo(5) & "'")
   aFormulaValue.Add CStr("'" & sListType & "'")
   aFormulaValue.Add optInc(0).Value
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

   sSql = "{RunsTable.RUNREF}='" & sPartNumber & "' " _
          & "AND {RunsTable.RUNNO}=" & Trim(cmbRun) & " "
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.SetDbTableConnection
   
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
'   SetMdiReportsize MDISect
'   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
'   MDISect.Crw.Formulas(1) = "Routing='" & sRout & "'"
'   MDISect.Crw.Formulas(2) = "RoutingRev='" & sRev & "'"
'   MDISect.Crw.Formulas(3) = "SoPon1='" & CustPo(1) & "'"
'   MDISect.Crw.Formulas(4) = "SoPon2='" & CustPo(2) & "'"
'   MDISect.Crw.Formulas(5) = "SoPon3='" & CustPo(3) & "'"
'   MDISect.Crw.Formulas(6) = "SoPon4='" & CustPo(4) & "'"
'   MDISect.Crw.Formulas(7) = "SoPon5='" & CustPo(5) & "'"
'   MDISect.Crw.Formulas(8) = "ListType='" & sListType & "'"
'
'   MDISect.Crw.ReportFileName = sReportPath & "awish01.rpt"
'   If optInc(0).Value = vbChecked Then
'      MDISect.Crw.SectionFormat(0) = "REPORTHDR.0.30;T;;;"
'   Else
'      MDISect.Crw.SectionFormat(0) = "REPORTHDR.0.30;F;;;"
'   End If
'   sSql = "{RunsTable.RUNREF}='" & sPartNumber & "' " _
'          & "AND {RunsTable.RUNNO}=" & Trim(cmbRun) & " "
'   MDISect.Crw.SelectionFormula = sSql
'   SetCrystalAction Me
   MouseCursor 0
   cmdCan.Enabled = True
   Exit Sub
   
Psh01:
   cmdCan.Enabled = True
   sProcName = "PrintSpecial"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
Psh02:
   DoModuleErrors Me
   
End Sub


'CJS Revised 9/6/06

Private Sub BuildAllocations()
   Dim RdoAlc As ADODB.Recordset
   Dim RdoIns As ADODB.Recordset
   Dim b As Byte
   Dim iList As Integer
   Dim iItem As Integer
   Dim sCust As String
   Dim sSon As String
   Dim sRev As String
   Dim sDate As String
   Dim sComt As String
   Dim sSoPoNum As String
   
   On Error Resume Next
   MouseCursor 13
   sSql = "TRUNCATE TABLE AwiSOAllocations"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "SELECT * FROM  AwiSOAllocations"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_KEYSET)
   'Requested by Janet Miller 8/28/01 apparently they can't allocate correctly
   sSql = "SELECT RAREF,RARUN,RASO,RASOITEM,RASOREV,ITSO," _
          & "ITNUMBER,ITREV,ITQTY FROM RnalTable,SoitTable " _
          & "WHERE (RASO=ITSO AND RASOITEM=ITNUMBER AND " _
          & "RASOREV=ITREV) AND (RAREF='" & Compress(cmbPrt) _
          & "' AND RARUN=" & Val(Trim(cmbRun)) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAlc, ES_STATIC)
   If bSqlRows Then
      With RdoAlc
         RdoIns.AddNew
         Do Until .EOF
            iList = iList + 1
            If iList > 5 Then Exit Do
            sSon = "" & str(!RASO)
            iItem = !RASOITEM
            sRev = "" & Trim(!RASOREV)
            sDate = ""
            sCust = ""
            b = GetSalesOrder(sCust, sSon, iItem, sRev, sDate, sComt, sSoPoNum)
            Select Case iList
               Case 1
                  RdoIns!AlcPart1 = "" & Trim(Compress(cmbPrt))
                  RdoIns!AlcCust1 = sCust
                  RdoIns!AlcSonm1 = sSon
                  RdoIns!AlcItem1 = Trim(str(iItem)) & sRev
                  RdoIns!AlcAqty1 = Format(!ITQTY, "#####0")
                  RdoIns!AlcDate1 = sDate
                  RdoIns!AlcComt1 = Trim(sComt)
                  RdoIns!AlcSoPo1 = Trim(sSoPoNum)
                  CustPo(1) = Trim(sSoPoNum)
                  
               Case 2
                  RdoIns!AlcPart2 = "" & Trim(Compress(cmbPrt))
                  RdoIns!AlcCust2 = sCust
                  RdoIns!AlcSonm2 = sSon
                  RdoIns!AlcItem2 = Trim(str(iItem)) & sRev
                  RdoIns!AlcAqty2 = Format(!ITQTY, "#####0")
                  RdoIns!AlcDate2 = sDate
                  RdoIns!AlcComt2 = Trim(sComt)
                  RdoIns!AlcSoPo2 = Trim(sSoPoNum)
                  CustPo(2) = Trim(sSoPoNum)
               Case 3
                  RdoIns!AlcPart3 = "" & Trim(Compress(cmbPrt))
                  RdoIns!AlcCust3 = sCust
                  RdoIns!AlcSonm3 = sSon
                  RdoIns!AlcItem3 = Trim(str(iItem)) & sRev
                  RdoIns!AlcAqty3 = Format(!ITQTY, "#####0")
                  RdoIns!AlcDate3 = sDate
                  RdoIns!AlcComt3 = Trim(sComt)
                  RdoIns!AlcSoPo3 = Trim(sSoPoNum)
                  CustPo(3) = Trim(sSoPoNum)
               Case 4
                  RdoIns!AlcPart4 = "" & Trim(Compress(cmbPrt))
                  RdoIns!AlcCust4 = sCust
                  RdoIns!AlcSonm4 = sSon
                  RdoIns!AlcItem4 = Trim(str(iItem)) & sRev
                  RdoIns!AlcAqty4 = Format(!ITQTY, "#####0")
                  RdoIns!AlcDate4 = sDate
                  RdoIns!AlcComt4 = Trim(sComt)
                  RdoIns!AlcSoPo4 = Trim(sSoPoNum)
                  CustPo(4) = Trim(sSoPoNum)
               Case Else
                  RdoIns!AlcPart5 = "" & Trim(Compress(cmbPrt))
                  RdoIns!AlcCust5 = sCust
                  RdoIns!AlcSonm5 = sSon
                  RdoIns!AlcItem5 = Trim(str(iItem)) & sRev
                  RdoIns!AlcAqty5 = Format(!ITQTY, "#####0")
                  RdoIns!AlcDate5 = sDate
                  RdoIns!AlcComt5 = Trim(sComt)
                  RdoIns!AlcSoPo5 = Trim(sSoPoNum)
                  CustPo(5) = Trim(sSoPoNum)
            End Select
            .MoveNext
         Loop
         RdoIns.Update
         ClearResultSet RdoAlc
      End With
   Else
      'One Record for the link
      RdoIns.AddNew
      RdoIns!AlcPart1 = "" & Compress(cmbPrt)
      RdoIns.Update
   End If
   MouseCursor 0
   Set RdoAlc = Nothing
   Set RdoIns = Nothing
   
End Sub

Private Function GetSalesOrder(Customer As String, SalesOrder As String, SoItem As Integer, SoRev As String, SoDate As String, _
                               Comments As String, PONUMBER As String) As Byte
   Dim RdoSon As ADODB.Recordset
   
   On Error Resume Next
   sSql = "SELECT SONUMBER,SOCUST,SOTYPE,SOPO,SOTEXT,ITSO,ITNUMBER,ITREV," _
          & "ITSCHED,ITCOMMENTS,CUREF,CUNICKNAME FROM SohdTable,SoitTable,CustTable " _
          & "WHERE (ITSO=" & Val(SalesOrder) & " AND ITNUMBER=" _
          & SoItem & " AND ITREV='" & SoRev & "' AND SONUMBER=ITSO " _
          & "AND SOCUST=CUREF)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         Customer = "" & Trim(!CUNICKNAME)
         SalesOrder = "" & Trim(!SOTYPE) & Trim(!SoText)
         SoItem = Format(!ITNUMBER, "###0")
         SoRev = "" & Trim(!ITREF)
         SoDate = Format(!ITSCHED, "mm/dd/yy")
         Comments = "" & Trim(!ITCOMMENTS)
         PONUMBER = "" & Trim(!SOPO)
         ClearResultSet RdoSon
      End With
   End If
   Set RdoSon = Nothing
   
End Function

Private Function GetMoRouting(Routing As String, Rev As String) As Byte
   Dim RdoRte As ADODB.Recordset
   Dim sRout As String
   sRout = Trim(lblTyp)
   sSql = "SELECT PARTREF,PAROUTING,RTREF,RTNUM,RTREV FROM PartTable," _
          & "RthdTable WHERE (PARTREF='" & Compress(cmbPrt) & "' " _
          & "AND PARTREF=RTREF)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
   If bSqlRows Then
      With RdoRte
         Routing = "" & Trim(!RTNUM)
         Rev = "" & Trim(!RTREV)
         ClearResultSet RdoRte
      End With
   End If
   If Routing = "" Then
      sSql = "SELECT RTEPART" & sRout & ",RTREF," _
             & "RTNUM,RTREV FROM ComnTable,RthdTable WHERE (COREF=1 " _
             & "AND RTEPART" & sRout & "=RTREF)"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte, ES_FORWARD)
      If bSqlRows Then
         With RdoRte
            Routing = "" & Trim(!RTNUM)
            Rev = "" & Trim(!RTREV)
            ClearResultSet RdoRte
         End With
      End If
   End If
   Set RdoRte = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getmorouting"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub BuildPoTable()
   Dim RdoPit As ADODB.Recordset
   Dim RdoIns As ADODB.Recordset
   Dim b As Byte
   Dim bb As Byte
   Dim sDesc As String
   Dim sUom As String
   Dim sPart As String
   Dim sVendor As String
   
   On Error Resume Next
   sSql = "TRUNCATE TABLE AwiPOAllocations"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "SELECT * FROM AwiPOAllocations"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_KEYSET)
   
   sSql = "SELECT PINUMBER,PIITEM,PIITEM,PIREV,PIPART,PIPDATE,PIADATE," _
          & "PIPQTY,PIAQTY,PILOT,PIRUNPART,PIRUNNO,PICOMT,PONUMBER,POVENDOR FROM " _
          & "PoitTable,PohdTable WHERE (PINUMBER=PONUMBER AND " _
          & "PIRUNPART='" & Compress(cmbPrt) & "' AND PIRUNNO=" & Val(cmbRun) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPit, ES_STATIC)
   If bSqlRows Then
      With RdoPit
         RdoIns.AddNew
         Do Until .EOF
            b = b + 1
            If b > 4 Then Exit Do
            sVendor = "" & Trim(!POVENDOR)
            bb = GetPoVendor(sVendor)
            sPart = "" & Trim(!PIPART)
            bb = GetPoPart(sPart, sDesc, sUom)
            On Error Resume Next
            Select Case b
               Case 1
                  RdoIns!PITRPart1 = Compress(cmbPrt)
                  RdoIns!PITPONum1 = Format$(!PINUMBER, "000000")
                  RdoIns!PITItem1 = !PIITEM
                  RdoIns!PITItRev1 = "" & Trim(!PIREV)
                  If !PIAQTY = 0 Then
                     RdoIns!PITPOQty1 = !PIPQTY
                     RdoIns!PITPODate1 = Format$(!PIPDATE, "mm/dd/yy")
                  Else
                     RdoIns!PITPOQty1 = !PIAQTY
                     RdoIns!PITPODate1 = Format$(!PIADATE, "mm/dd/yy") & "*"
                  End If
                  RdoIns!PITPOPart1 = sPart
                  RdoIns!PITPOUom1 = sUom
                  RdoIns!PITPODesc1 = "" & Trim(sDesc)
                  RdoIns!PITPOCmt1 = "" & Trim(!PICOMT)
                  RdoIns!PITPOVend1 = sVendor
               Case 2
                  RdoIns!PITRPart2 = Compress(cmbPrt)
                  RdoIns!PITPONum2 = Format$(!PINUMBER, "000000")
                  RdoIns!PITItem2 = !PIITEM
                  RdoIns!PITItRev2 = "" & Trim(!PIREV)
                  If !PIAQTY = 0 Then
                     RdoIns!PITPOQty2 = !PIPQTY
                     RdoIns!PITPODate2 = Format$(!PIPDATE, "mm/dd/yy")
                  Else
                     RdoIns!PITPOQty2 = !PIAQTY
                     RdoIns!PITPODate2 = Format$(!PIADATE, "mm/dd/yy") & "*"
                  End If
                  RdoIns!PITPOPart2 = sPart
                  RdoIns!PITPOUom2 = sUom
                  RdoIns!PITPODesc2 = "" & Trim(sDesc)
                  RdoIns!PITPOCmt2 = "" & Trim(!PICOMT)
                  RdoIns!PITPOVend2 = sVendor
               Case 3
                  RdoIns!PITRPart3 = Compress(cmbPrt)
                  RdoIns!PITPONum3 = Format$(!PINUMBER, "000000")
                  RdoIns!PITItem3 = !PIITEM
                  RdoIns!PITItRev3 = "" & Trim(!PIREV)
                  If !PIAQTY = 0 Then
                     RdoIns!PITPOQty3 = !PIPQTY
                     RdoIns!PITPODate3 = Format$(!PIPDATE, "mm/dd/yy")
                  Else
                     RdoIns!PITPOQty3 = !PIAQTY
                     RdoIns!PITPODate3 = Format$(!PIADATE, "mm/dd/yy") & "*"
                  End If
                  RdoIns!PITPOPart3 = sPart
                  RdoIns!PITPOUom3 = sUom
                  RdoIns!PITPODesc3 = "" & Trim(sDesc)
                  RdoIns!PITPOCmt3 = "" & Trim(!PICOMT)
                  RdoIns!PITPOVend3 = sVendor
               Case Else
                  RdoIns!PITRPart4 = Compress(cmbPrt)
                  RdoIns!PITPONum4 = Format$(!PINUMBER, "000000")
                  RdoIns!PITItem4 = !PIITEM
                  RdoIns!PITItRev4 = "" & Trim(!PIREV)
                  If !PIAQTY = 0 Then
                     RdoIns!PITPOQty4 = !PIPQTY
                     RdoIns!PITPODate4 = Format$(!PIPDATE, "mm/dd/yy")
                  Else
                     RdoIns!PITPOQty4 = !PIAQTY
                     RdoIns!PITPODate4 = Format$(!PIADATE, "mm/dd/yy") & "*"
                  End If
                  RdoIns!PITPOPart4 = sPart
                  RdoIns!PITPOUom4 = sUom
                  RdoIns!PITPODesc4 = "" & Trim(sDesc)
                  RdoIns!PITPOCmt4 = "" & Trim(!PICOMT)
                  RdoIns!PITPOVend4 = sVendor
            End Select
            .MoveNext
         Loop
         RdoIns.Update
         ClearResultSet RdoPit
      End With
   Else
      'one dummy for the link
      RdoIns.AddNew
      RdoIns!PITRPart1 = Compress(cmbPrt)
      RdoIns.Update
   End If
   Set RdoPit = Nothing
   Set RdoIns = Nothing
   Exit Sub
   
   
End Sub

Private Function GetPoVendor(VENDOR As String) As Byte
   Dim RdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetVendorBasics '" & VENDOR & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
   If bSqlRows Then
      With RdoVnd
         VENDOR = "" & Trim(!VENICKNAME)
         ClearResultSet RdoVnd
      End With
   End If
   Set RdoVnd = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpovendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPoPart(part As String, Desc As String, Units As String) As Byte
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS FROM PartTable " _
          & "WHERE PARTREF='" & part & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         part = "" & Trim(!PartNum)
         Desc = "" & Trim(!PADESC)
         Units = "" & Trim(!PAUNITS)
         ClearResultSet RdoPrt
      End With
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpopart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function





'CJS 9/6/06

Private Sub CreateSQLSOAllocations()
   On Error Resume Next
   sSql = "SELECT AlcPart1 FROM AwiSOAllocations"
   clsADOCon.ExecuteSQL sSql
   If Err > 0 Then
      sSql = "CREATE TABLE AwiSOAllocations (" _
             & "AlcPart1 CHAR(30) NOT NULL," _
             & "AlcCust1 CHAR(10) NULL DEFAULT('')," _
             & "AlcSonm1 CHAR(6) NULL DEFAULT('')," _
             & "AlcItem1 CHAR(10) NULL DEFAULT('')," _
             & "AlcAQty1 smallmoney NULL DEFAULT(0)," _
             & "AlcDate1 CHAR(8) NULL DEFAULT('')," _
             & "AlcComt1 CHAR(255) NULL DEFAULT('')," _
             & "AlcSopo1 CHAR(20) NULL DEFAULT('')," _
             & "AlcPart2 CHAR(30) NULL DEFAULT('')," _
             & "AlcCust2 CHAR(10) NULL DEFAULT('')," _
             & "AlcSonm2 CHAR(6) NULL DEFAULT('')," _
             & "AlcItem2 CHAR(10) NULL DEFAULT('')," _
             & "AlcAQty2 smallmoney NULL DEFAULT(0)," _
             & "AlcDate2 CHAR(8) NULL DEFAULT('')," _
             & "AlcComt2 CHAR(255) NULL DEFAULT('')," _
             & "AlcSopo2 CHAR(20) NULL DEFAULT(''),"
      sSql = sSql _
             & "AlcPart3 CHAR(30) NULL DEFAULT('')," _
             & "AlcCust3 CHAR(10) NULL DEFAULT('')," _
             & "AlcSonm3 CHAR(6) NULL DEFAULT('')," _
             & "AlcItem3 CHAR(10) NULL DEFAULT('')," _
             & "AlcAQty3 smallmoney NULL DEFAULT(0)," _
             & "AlcDate3 CHAR(8) NULL DEFAULT('')," _
             & "AlcComt3 CHAR(255) NULL DEFAULT('')," _
             & "AlcSopo3 CHAR(20) NULL DEFAULT('')," _
             & "AlcPart4 CHAR(30) NULL DEFAULT('')," _
             & "AlcCust4 CHAR(10) NULL DEFAULT('')," _
             & "AlcSonm4 CHAR(6) NULL DEFAULT('')," _
             & "AlcItem4 CHAR(10) NULL DEFAULT('')," _
             & "AlcAQty4 smallmoney NULL DEFAULT(0)," _
             & "AlcDate4 CHAR(8) NULL DEFAULT('')," _
             & "AlcComt4 CHAR(255) NULL DEFAULT('')," _
             & "AlcSopo4 CHAR(20) NULL DEFAULT(''),"
      sSql = sSql _
             & "AlcPart5 CHAR(30) NULL DEFAULT('')," _
             & "AlcCust5 CHAR(10) NULL DEFAULT('')," _
             & "AlcSonm5 CHAR(6) NULL DEFAULT('')," _
             & "AlcItem5 CHAR(10) NULL DEFAULT('')," _
             & "AlcAQty5 smallmoney NULL DEFAULT(0)," _
             & "AlcDate5 CHAR(8) NULL DEFAULT('')," _
             & "AlcComt5 CHAR(255) NULL DEFAULT('')," _
             & "AlcSopo5 CHAR(20) NULL DEFAULT(''))"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE AwiSOAllocations ADD Constraint PK_AwiSOAllocations_ALLOCREF PRIMARY KEY CLUSTERED (AlcPart1) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
   Else
      sSql = "TRUNCATE TABLE AwiSOAllocations"
      clsADOCon.ExecuteSQL sSql
      
   End If
End Sub

'CJS 9/6/06

Private Sub CreateSQLPOAllocations()
   On Error Resume Next
   sSql = "SELECT PITRPart1 FROM AwiPOAllocations"
   clsADOCon.ExecuteSQL sSql
   If Err > 0 Then
      Err.Clear
      sSql = "CREATE TABLE AwiPOAllocations (" _
             & "PITRPart1 CHAR(30) NOT NULL," _
             & "PITPONum1 CHAR(6) NULL DEFAULT('')," _
             & "PITItem1 smallint NULL DEFAULT(0)," _
             & "PITItRev1 CHAR(2) NULL DEFAULT('')," _
             & "PITPOUom1 CHAR(2) NULL DEFAULT('')," _
             & "PITPOQty1 smallmoney NULL DEFAULT(0)," _
             & "PITPODate1 CHAR(9) NULL DEFAULT('')," _
             & "PITPOPart1 CHAR(30) NULL DEFAULT('')," _
             & "PITPODesc1 CHAR(30) NULL DEFAULT('')," _
             & "PITPOCmt1 CHAR(255) NULL DEFAULT('')," _
             & "PITPOVend1 CHAR(10) NULL DEFAULT('')," _
             & "PITRPart2 CHAR(30) NULL DEFAULT('')," _
             & "PITPONum2 CHAR(6) NULL DEFAULT('')," _
             & "PITItem2 smallint NULL DEFAULT(0)," _
             & "PITItRev2 CHAR(2) NULL DEFAULT('')," _
             & "PITPOUom2 CHAR(2) NULL DEFAULT('')," _
             & "PITPOQty2 smallmoney NULL DEFAULT(0)," _
             & "PITPODate2 CHAR(9) NULL DEFAULT('')," _
             & "PITPOPart2 CHAR(30) NULL DEFAULT('')," _
             & "PITPODesc2 CHAR(30) NULL DEFAULT('')," _
             & "PITPOCmt2 CHAR(255) NULL DEFAULT(''),"
      sSql = sSql _
             & "PITPOVend2 CHAR(10) NULL DEFAULT('')," _
             & "PITRPart3 CHAR(30) NULL DEFAULT('')," _
             & "PITPONum3 CHAR(6) NULL DEFAULT('')," _
             & "PITItem3 smallint NULL DEFAULT(0)," _
             & "PITItRev3 CHAR(2) NULL DEFAULT('')," _
             & "PITPOUom3 CHAR(2) NULL DEFAULT('')," _
             & "PITPOQty3 smallmoney NULL DEFAULT(0)," _
             & "PITPODate3 CHAR(9) NULL DEFAULT('')," _
             & "PITPOPart3 CHAR(30) NULL DEFAULT('')," _
             & "PITPODesc3 CHAR(30) NULL DEFAULT('')," _
             & "PITPOCmt3 CHAR(255) NULL DEFAULT('')," _
             & "PITPOVend3 CHAR(10) NULL DEFAULT('')," _
             & "PITRPart4 CHAR(30) NULL DEFAULT('')," _
             & "PITPONum4 CHAR(6) NULL DEFAULT('')," _
             & "PITItem4 smallint NULL DEFAULT(0)," _
             & "PITItRev4 CHAR(2) NULL DEFAULT('')," _
             & "PITPOUom4 CHAR(2) NULL DEFAULT('')," _
             & "PITPOQty4 smallmoney NULL DEFAULT(0)," _
             & "PITPODate4 CHAR(9) NULL DEFAULT('')," _
             & "PITPOPart4 CHAR(30) NULL DEFAULT('')," _
             & "PITPODesc4 CHAR(30) NULL DEFAULT('')," _
             & "PITPOCmt4 CHAR(255) NULL DEFAULT('')," _
             & "PITPOVend4 CHAR(10) NULL DEFAULT(''))"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE AwiPOAllocations ADD Constraint PK_AwiPOAllocations_ALLOCREF PRIMARY KEY CLUSTERED (PITRPart1) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
   Else
      sSql = "TRUNCATE TABLE AwiPOAllocations"
      clsADOCon.ExecuteSQL sSql
   End If
   
End Sub

'CJS 9/6/06

Private Sub CreateSQLPLAllocations()
   On Error Resume Next
   sSql = "SELECT PLSPart1 FROM AwiPLAllocations"
   clsADOCon.ExecuteSQL sSql
   If Err > 0 Then
      Err.Clear
      sSql = "CREATE TABLE AwiPLAllocations (" _
             & "PLSRunPart CHAR(30) NOT NULL," _
             & "PLSPart1 CHAR(30) NULL DEFAULT('')," _
             & "PLSDesc1 CHAR(30) NULL DEFAULT('')," _
             & "PLSADate1 CHAR(8) NULL DEFAULT('')," _
             & "PLSAQty1 smallmoney NULL DEFAULT(0)," _
             & "PLSPQty1 smallmoney NULL DEFAULT(0)," _
             & "PLSUom1 CHAR(2) NULL DEFAULT('')," _
             & "PLSALoc1 CHAR(4) NULL DEFAULT('')," _
             & "PLSCom1 CHAR(255) NULL DEFAULT('')," _
             & "PLSPart2 CHAR(30) NULL DEFAULT('')," _
             & "PLSDesc2 CHAR(30) NULL DEFAULT('')," _
             & "PLSADate2 CHAR(8) NULL DEFAULT('')," _
             & "PLSAQty2 smallmoney NULL DEFAULT(0)," _
             & "PLSPQty2 smallmoney NULL DEFAULT(0)," _
             & "PLSUom2 CHAR(2) NULL DEFAULT('')," _
             & "PLSALoc2 CHAR(4) NULL DEFAULT('')," _
             & "PLSCom2 CHAR(255) NULL DEFAULT(''),"
      sSql = sSql _
             & "PLSPart3 CHAR(30) NULL DEFAULT('')," _
             & "PLSDesc3 CHAR(30) NULL DEFAULT('')," _
             & "PLSADate3 CHAR(8) NULL DEFAULT('')," _
             & "PLSAQty3 smallmoney NULL DEFAULT(0)," _
             & "PLSPQty3 smallmoney NULL DEFAULT(0)," _
             & "PLSUom3 CHAR(2) NULL DEFAULT('')," _
             & "PLSALoc3 CHAR(4) NULL DEFAULT('')," _
             & "PLSCom3 CHAR(255) NULL DEFAULT('')," _
             & "PLSPart4 CHAR(30) NULL DEFAULT('')," _
             & "PLSDesc4 CHAR(30) NULL DEFAULT('')," _
             & "PLSADate4 CHAR(8) NULL DEFAULT('')," _
             & "PLSAQty4 smallmoney NULL DEFAULT(0)," _
             & "PLSPQty4 smallmoney NULL DEFAULT(0)," _
             & "PLSUom4 CHAR(2) NULL DEFAULT('')," _
             & "PLSALoc4 CHAR(4) NULL DEFAULT('')," _
             & "PLSCom4 CHAR(255) NULL DEFAULT(''))"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "ALTER TABLE AwiPLAllocations ADD Constraint PK_AwiPLAllocations_ALLOCREF PRIMARY KEY CLUSTERED (PLSRunPart) " _
             & "WITH FILLFACTOR=80 "
      clsADOCon.ExecuteSQL sSql
   Else
      sSql = "TRUNCATE TABLE AwiPLAllocations"
      clsADOCon.ExecuteSQL sSql
   End If
   
   
End Sub
