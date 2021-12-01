VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Income Statment With Percentages (Report)"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbact 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   24
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   3360
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   3600
      Width           =   735
   End
   Begin VB.CheckBox chkcon 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtLvl 
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Tag             =   "1"
      Top             =   2880
      Width           =   285
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   2280
      TabIndex        =   12
      Tag             =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Tag             =   "4"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   120
         Picture         =   "diaGLp07a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Display The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   720
         Picture         =   "diaGLp07a.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.ComboBox txtYearBeg 
      Height          =   315
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "4"
      ToolTipText     =   "Enter New Team Member  (15 Char) Or Select From List"
      Top             =   720
      Width           =   1095
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   1
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
      PictureUp       =   "diaGLp07a.frx":0308
      PictureDn       =   "diaGLp07a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4590
      FormDesignWidth =   6465
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaGLp07a.frx":0594
      PictureDn       =   "diaGLp07a.frx":06DA
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   25
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ratio Master Account Is"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inactive Accounts"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Accounts W/O Divisions"
      Height          =   405
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consolidated"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Divisionalized Reports Only)"
      Height          =   285
      Index           =   8
      Left            =   3960
      TabIndex        =   19
      Top             =   3600
      Width           =   2145
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(9 For All)"
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   15
      Top             =   2880
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through Detail Level"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1545
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Ending Date"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Beginning Date"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Year Beginning Date"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "diaGLp07a"
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
   '
   ' diaGLp07a - Income Statement
   '
   ' Notes: Used the income statement with percentages form and report as a base.
   '
   ' Created:  9/30/01 (nth)
   ' Revisions:
   '
   '*************************************************************************************
   
   Option Explicit
   
   Dim bOnLoad As Byte
   Dim vAccounts(10, 4) As Variant
   Dim iTotal As Integer
   Dim iInActive As Integer
   Dim iCurType As Integer
   Dim iFsLevel As Integer
   Dim iLevel As Integer
   Dim sPrevYearBeg As String
   Dim sPrevYearEnd As String
   Dim DbAct As Recordset 'Jet
   Dim cRatio1 As Currency
   Dim cRatio2 As Currency
   Dim cRatio3 As Currency
   Dim AdoQry1 As ADODB.Command
   Dim AdoParameter1 As ADODB.Parameter
   Dim AdoParameter2 As ADODB.Parameter
   
   Private txtKeyPress() As New EsiKeyBd
   Private txtGotFocus() As New EsiKeyBd
   Private txtKeyDown() As New EsiKeyBd
   
   Private Sub cmbAct_Click()
      lblDsc = UpdateActDesc(cmbact)
   End Sub
   
   Private Sub cmbAct_LostFocus()
      lblDsc = UpdateActDesc(cmbact)
   End Sub
   
   Private Sub cmdCan_Click()
      Unload Me
   End Sub
   
   Private Sub cmdHlp_Click(Value As Integer)
      If cmdHlp Then
         MouseCursor 13
         SelectHelpTopic Me, "No Subject Help"
         cmdHlp = False
         MouseCursor 0
      End If
      
   End Sub
   
   Private Sub Form_Activate()
      MdiSect.lblBotPanel = Caption
      If bOnLoad Then
         FillCombo
         CreateActTable
         bOnLoad = False
      End If
      MouseCursor 0
   End Sub
   
   Private Sub Form_Load()
      FormLoad Me
      FormatControls
      ReopenJet
      sCurrForm = Caption
      txtYearBeg = "01/01/" & Format(Now, "yy")
      txtBeg = Format(Now, "mm/01/yy")
      ' Assign value to txtend
      GetMonthEnd txtBeg
      sSql = "SELECT JIACCOUNT, SUM(JIDEB)-SUM(JICRD) " _
             & "FROM GjitTable INNER JOIN " _
             & "GjhdTable ON GjitTable.JINAME = GjhdTable.GJNAME " _
             & "WHERE GjhdTable.GJPOSTED=1 " _
             & " AND GjhdTable.GJPOST >=?" _
             & " AND GjhdTable.GJPOST <=?" _
             & " GROUP BY JIACCOUNT"
      Set AdoQry1 = New ADODB.Command
      AdoQry1.CommandText = sSql
      
      Set AdoParameter1 = New ADODB.Parameter
      AdoParameter1.Type = adDate
      Set AdoParameter2 = New ADODB.Parameter
      AdoParameter2.Type = adDate
      AdoQry1.parameters.Append AdoParameter1
      AdoQry1.parameters.Append AdoParameter2
      
      chkcon.Value = 1 ' temporary
      bOnLoad = True
   End Sub
   
   Private Sub Form_Resize()
      Refresh
   End Sub
   
   Private Sub Form_Unload(Cancel As Integer)
      FormUnload
      On Error Resume Next
      'JetDb.Execute "DROP TABLE ActrpTable"
      'JetDb.Execute "DROP TABLE Current"
      'JetDb.Execute "DROP TABLE Previous"
      'JetDb.Execute "DROP TABLE YTD"
      Set AdoParameter1 = Nothing
      Set AdoParameter2 = Nothing
      Set AdoQry1 = Nothing
      
      Set diaGLp07a = Nothing
   End Sub
   
   Private Sub FormatControls()
      Dim b As Byte
      b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   End Sub
   
   Private Sub FillCombo()
      Dim rdoAct As ADODB.Recordset
      sSql = "SELECT GLACCTNO FROM GlacTable"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct)
      If bSqlRows Then
         With rdoAct
            While Not .EOF
               AddComboStr cmbact.hwnd, "" & Trim(!GLACCTNO)
               .MoveNext
            Wend
         End With
      End If
      Set rdoAct = Nothing
      cmbact.ListIndex = 0
      Exit Sub
DiaErr1:
      sProcName = "fillcombo"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Public Sub BuildAccounts()
      Dim i As Integer
      Dim RdoGlm As ADODB.Recordset
      Dim RdoAct1 As ADODB.Recordset
      Dim RdoAct2 As ADODB.Recordset
      Dim RdoAct3 As ADODB.Recordset
      Dim DbBal1 As Recordset
      Dim DbBal2 As Recordset
      Dim DbBal3 As Recordset
      Dim sAccount As String
      Dim sRatioAcct As String
      
      MouseCursor 13
      On Error GoTo DiaErr1
      If Trim(txtLvl) = "" Then txtLvl = 9
      iLevel = Val(txtLvl)
      
      ' Build income statement account structure
      sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
      If bSqlRows Then
         With RdoGlm
            i = 4
            vAccounts(i, 0) = "" & Trim(!COINCMREF)
            vAccounts(i, 1) = "" & Trim(!COINCMACCT)
            vAccounts(i, 2) = "" & Trim(!COINCMDESC)
            vAccounts(i, 3) = Format(!COINCMTYPE, "0")
            
            i = 5
            vAccounts(i, 0) = "" & Trim(!COCOGSREF)
            vAccounts(i, 1) = "" & Trim(!COCOGSACCT)
            vAccounts(i, 2) = "" & Trim(!COCOGSDESC)
            vAccounts(i, 3) = Format(!COCOGSTYPE, "0")
            
            i = 6
            vAccounts(i, 0) = "" & Trim(!COEXPNREF)
            vAccounts(i, 1) = "" & Trim(!COEXPNACCT)
            vAccounts(i, 2) = "" & Trim(!COEXPNDESC)
            vAccounts(i, 3) = Format(!COEXPNTYPE, "0")
            
            i = 7
            vAccounts(i, 0) = "" & Trim(!COOINCREF)
            vAccounts(i, 1) = "" & Trim(!COOINCACCT)
            vAccounts(i, 2) = "" & Trim(!COOINCDESC)
            vAccounts(i, 3) = Format(!COOINCTYPE, "0")
            
            i = 8
            vAccounts(i, 0) = "" & Trim(!COOEXPREF)
            vAccounts(i, 1) = "" & Trim(!COOEXPACCT)
            vAccounts(i, 2) = "" & Trim(!COOEXPDESC)
            vAccounts(i, 3) = Format(!COOEXPTYPE, "0")
            
            i = 9
            vAccounts(i, 0) = "" & Trim(!COFDTXREF)
            vAccounts(i, 1) = "" & Trim(!COFDTXACCT)
            vAccounts(i, 2) = "" & Trim(!COFDTXDESC)
            vAccounts(i, 3) = Format(!COFDTXTYPE, "0")
         End With
      End If
      iTotal = i
      Set RdoGlm = Nothing
      
      'Clear temp jet tables
      sSql = "DELETE FROM ActrpTable"
      JetDb.Execute sSql
      sSql = "DELETE FROM Current"
      JetDb.Execute sSql
      sSql = "DELETE FROM YTD"
      JetDb.Execute sSql
      sSql = "DELETE FROM Previous"
      JetDb.Execute sSql
      
      ' Populate the finacial statement layout JET table
      Set DbAct = JetDb.OpenRecordset("ActrpTable", dbOpenDynaset)
      For i = 4 To iTotal
         iCurType = i
         FillLevel1 Format(vAccounts(i, 0))
      Next
      DbAct.Close
      
      sPrevYearBeg = DateAdd("yyyy", -1, txtYearBeg)
      sPrevYearEnd = DateAdd("d", -1, txtYearBeg)
      
      ' Fill current account balance table
      Set DbBal1 = JetDb.OpenRecordset("Current", dbOpenDynaset)
      AdoQry1.parameters(0).Value = "" & txtBeg
      AdoQry1.parameters(1).Value = "" & txtEnd
      bSqlRows = clsADOCon.GetQuerySet(RdoAct1, AdoQry1, ES_FORWARD)
      If bSqlRows Then
         With RdoAct1
            While Not .EOF
               DbBal1.AddNew
               DbBal1!ACCTREF = "" & Trim(!JIACCOUNT)
               DbBal1!ACCTBAL = .Fields(1)
               DbBal1.Update
               .MoveNext
            Wend
         End With
      End If
      Set RdoAct1 = Nothing
      DbBal1.Close
      
      ' Fill YTD account balance table
      Set DbBal2 = JetDb.OpenRecordset("YTD", dbOpenDynaset)
      AdoQry1.parameters(0).Value = "" & txtYearBeg
      AdoQry1.parameters(1).Value = "" & txtEnd
      bSqlRows = clsADOCon.GetQuerySet(RdoAct2, AdoQry1, ES_FORWARD)
      If bSqlRows Then
         With RdoAct2
            While Not .EOF
               DbBal2.AddNew
               DbBal2!ACCTREF = "" & Trim(!JIACCOUNT)
               DbBal2!ACCTBAL = .Fields(1)
               DbBal2.Update
               .MoveNext
            Wend
         End With
      End If
      Set RdoAct2 = Nothing
      DbBal2.Close
      
      ' Fill previous account balance table
      Set DbBal3 = JetDb.OpenRecordset("Previous", dbOpenDynaset)
      AdoQry1.parameters(0).Value = sPrevYearBeg
      AdoQry1.parameters(1).Value = sPrevYearEnd
      bSqlRows = clsADOCon.GetQuerySet(RdoAct3, AdoQry1, ES_FORWARD)
      If bSqlRows Then
         With RdoAct3
            While Not .EOF
               DbBal3.AddNew
               DbBal3!ACCTREF = "" & Trim(!JIACCOUNT)
               DbBal3!ACCTBAL = .Fields(1)
               DbBal3.Update
               .MoveNext
            Wend
         End With
      End If
      Set RdoAct3 = Nothing
      DbBal3.Close
      
      ReopenJet
      PrintReport
      Exit Sub
      
DiaErr1:
      sProcName = "buildaccou"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Public Sub FillLevel1(sMaster As String)
      Dim i As Integer
      Dim RdoAct1 As ADODB.Recordset
      Dim iRemCurType As Integer
      Dim iRemFsLevel As Integer
      
      On Error GoTo DiaErr1
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct1)
      If bSqlRows Then
         With RdoAct1
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 1
                  !Act02 = "" & Trim(RdoAct1!GLACCTREF)
                  !Act03 = String$(2, Chr$(160)) & "" & Trim(RdoAct1!GLACCTNO)
                  !Act04 = String$(2, Chr$(160)) & "" & Trim(RdoAct1!GLDESCR)
                  !Act05 = RdoAct1!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct1!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               iRemCurType = iCurType
               iRemFsLevel = iFsLevel
               
               If iLevel > 1 Then FillLevel2 Trim(!GLACCTREF)
               If iRemFsLevel = 1 Then
                  With DbAct
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 1
                     !Act02 = "" & Trim(RdoAct1!GLACCTREF)
                     !Act04 = String$(2, Chr$(160)) & "Total " & Trim(RdoAct1!GLDESCR)
                     !Act05 = RdoAct1!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 1
                     !Act02 = "" & Trim(RdoAct1!GLACCTREF)
                     !Act04 = ""
                     !Act05 = RdoAct1!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                  End With
               End If
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct1 = Nothing
      
      
      
      Exit Sub
DiaErr1:
      sProcName = "filllevel1"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Public Sub FillLevel2(sMaster As String)
      Dim i As Integer
      Dim RdoAct2 As ADODB.Recordset
      Dim iRemFsLevel As Integer
      Dim iRemCurType As Integer
      
      On Error GoTo DiaErr1
      
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct2)
      If bSqlRows Then
         With RdoAct2
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 2
                  !Act02 = "" & Trim(RdoAct2!GLACCTREF)
                  !Act03 = String$(6, Chr$(160)) & "" & Trim(RdoAct2!GLACCTNO)
                  !Act04 = String$(6, Chr$(160)) & "" & Trim(RdoAct2!GLDESCR)
                  !Act05 = RdoAct2!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct2!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               iRemCurType = iCurType
               iRemFsLevel = iFsLevel
               If iLevel > 2 Then FillLevel3 Trim(!GLACCTREF)
               If iRemFsLevel = 2 Then
                  With DbAct
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 2
                     !Act02 = "" & Trim(RdoAct2!GLACCTREF)
                     
                     !Act04 = String$(6, Chr$(160)) & "Total " & Trim(RdoAct2!GLDESCR)
                     !Act05 = RdoAct2!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 2
                     !Act02 = "" & Trim(RdoAct2!GLACCTREF)
                     !Act04 = ""
                     !Act05 = RdoAct2!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                  End With
               End If
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct2 = Nothing
      Exit Sub
DiaErr1:
      sProcName = "filllevel2"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Public Sub FillLevel3(sMaster As String)
      Dim i As Integer
      Dim RdoAct3 As ADODB.Recordset
      Dim iRemFsLevel As Integer
      Dim iRemCurType As Integer
      
      On Error GoTo DiaErr1
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct3)
      If bSqlRows Then
         With RdoAct3
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 3
                  !Act02 = "" & Trim(RdoAct3!GLACCTREF)
                  !Act03 = String$(8, Chr$(160)) & "" & Trim(RdoAct3!GLACCTNO)
                  !Act04 = String$(8, Chr$(160)) & "" & Trim(RdoAct3!GLDESCR)
                  !Act05 = RdoAct3!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct3!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               iRemCurType = iCurType
               iRemFsLevel = iFsLevel
               If iLevel > 3 Then FillLevel4 Trim(!GLACCTREF)
               If iRemFsLevel = 3 Then
                  With DbAct
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 3
                     !Act02 = "" & Trim(RdoAct3!GLACCTREF)
                     
                     !Act04 = String$(8, Chr$(160)) & "Total " & Trim(RdoAct3!GLDESCR)
                     !Act05 = RdoAct3!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 3
                     !Act02 = "" & Trim(RdoAct3!GLACCTREF)
                     !Act04 = ""
                     !Act05 = RdoAct3!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                  End With
               End If
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct3 = Nothing
      Exit Sub
DiaErr1:
      sProcName = "filllevel3"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Public Sub FillLevel4(sMaster As String)
      Dim i As Integer
      Dim RdoAct4 As ADODB.Recordset
      Dim iRemFsLevel As Integer
      Dim iRemCurType As Integer
      
      On Error GoTo DiaErr1
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct4)
      If bSqlRows Then
         With RdoAct4
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 4
                  !Act02 = "" & Trim(RdoAct4!GLACCTREF)
                  !Act03 = String$(10, Chr$(160)) & "" & Trim(RdoAct4!GLACCTNO)
                  !Act04 = String$(10, Chr$(160)) & "" & Trim(RdoAct4!GLDESCR)
                  !Act05 = RdoAct4!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct4!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               iRemCurType = iCurType
               iRemFsLevel = iFsLevel
               If iLevel > 4 Then FillLevel5 Trim(!GLACCTREF)
               If iRemFsLevel = 4 Then
                  With DbAct
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 4
                     !Act02 = "" & Trim(RdoAct4!GLACCTREF)
                     
                     !Act04 = String$(10, Chr$(160)) & "Total " & Trim(RdoAct4!GLDESCR)
                     !Act05 = RdoAct4!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 4
                     !Act02 = "" & Trim(RdoAct4!GLACCTREF)
                     !Act04 = ""
                     !Act05 = RdoAct4!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                  End With
               End If
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct4 = Nothing
      Exit Sub
DiaErr1:
      sProcName = "filllevel4"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Public Sub FillLevel5(sMaster As String)
      Dim i As Integer
      Dim RdoAct5 As ADODB.Recordset
      Dim iRemFsLevel As Integer
      Dim iRemCurType As Integer
      
      On Error GoTo DiaErr1
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct5)
      If bSqlRows Then
         With RdoAct5
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 5
                  !Act02 = "" & Trim(RdoAct5!GLACCTREF)
                  !Act03 = String$(12, Chr$(160)) & "" & Trim(RdoAct5!GLACCTNO)
                  !Act04 = String$(12, Chr$(160)) & "" & Trim(RdoAct5!GLDESCR)
                  !Act05 = RdoAct5!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct5!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               iRemCurType = iCurType
               iRemFsLevel = iFsLevel
               If iLevel > 5 Then FillLevel6 Trim(!GLACCTREF)
               If iRemFsLevel = 5 Then
                  With DbAct
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 5
                     !Act02 = "" & Trim(RdoAct5!GLACCTREF)
                     
                     !Act04 = String$(12, Chr$(160)) & "Total " & Trim(RdoAct5!GLDESCR)
                     !Act05 = RdoAct5!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 5
                     !Act02 = "" & Trim(RdoAct5!GLACCTREF)
                     !Act04 = ""
                     !Act05 = RdoAct5!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                  End With
               End If
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct5 = Nothing
      Exit Sub
DiaErr1:
      sProcName = "filllevel5"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
      
   End Sub
   
   Public Sub FillLevel6(sMaster As String)
      Dim i As Integer
      Dim RdoAct6 As ADODB.Recordset
      Dim RdoBal6 As ADODB.Recordset
      Dim iRemFsLevel As Integer
      Dim iRemCurType As Integer
      
      On Error GoTo DiaErr1
      
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct6)
      If bSqlRows Then
         With RdoAct6
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 6
                  !Act02 = "" & Trim(RdoAct6!GLACCTREF)
                  !Act03 = String$(14, Chr$(160)) & "" & Trim(RdoAct6!GLACCTNO)
                  !Act04 = String$(14, Chr$(160)) & "" & Trim(RdoAct6!GLDESCR)
                  !Act05 = RdoAct6!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct6!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               iRemCurType = iCurType
               iRemFsLevel = iFsLevel
               If iLevel > 6 Then FillLevel7 Trim(!GLACCTREF)
               If iRemFsLevel = 6 Then
                  With DbAct
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 6
                     !Act02 = "" & Trim(RdoAct6!GLACCTREF)
                     
                     !Act04 = String$(14, Chr$(160)) & "Total " & Trim(RdoAct6!GLDESCR)
                     !Act05 = RdoAct6!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 6
                     !Act02 = "" & Trim(RdoAct6!GLACCTREF)
                     !Act04 = ""
                     !Act05 = RdoAct6!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                  End With
               End If
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct6 = Nothing
      Exit Sub
DiaErr1:
      sProcName = "filllevel6"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
      
   End Sub
   
   Public Sub FillLevel7(sMaster As String)
      Dim i As Integer
      Dim RdoAct7 As ADODB.Recordset
      Dim iRemFsLevel As Integer
      Dim iRemCurType As Integer
      
      On Error GoTo DiaErr1
      
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct7)
      If bSqlRows Then
         With RdoAct7
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 7
                  !Act02 = "" & Trim(RdoAct7!GLACCTREF)
                  !Act03 = String$(16, Chr$(160)) & "" & Trim(RdoAct7!GLACCTNO)
                  !Act04 = String$(16, Chr$(160)) & "" & Trim(RdoAct7!GLDESCR)
                  !Act05 = RdoAct7!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct7!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               iRemCurType = iCurType
               iRemFsLevel = iFsLevel
               If iLevel > 7 Then FillLevel8 Trim(!GLACCTREF)
               If iRemFsLevel = 7 Then
                  With DbAct
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 7
                     !Act02 = "" & Trim(RdoAct7!GLACCTREF)
                     
                     !Act04 = String$(16, Chr$(160)) & "Total " & Trim(RdoAct7!GLDESCR)
                     !Act05 = RdoAct7!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 7
                     !Act02 = "" & Trim(RdoAct7!GLACCTREF)
                     !Act04 = ""
                     !Act05 = RdoAct7!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                  End With
               End If
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct7 = Nothing
      Exit Sub
DiaErr1:
      sProcName = "filllevel7"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
      
   End Sub
   
   Public Sub FillLevel8(sMaster As String)
      Dim i As Integer
      Dim RdoAct8 As ADODB.Recordset
      Dim iRemFsLevel As Integer
      Dim iRemCurType As Integer
      
      On Error GoTo DiaErr1
      
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct8)
      If bSqlRows Then
         With RdoAct8
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 8
                  !Act02 = "" & Trim(RdoAct8!GLACCTREF)
                  !Act03 = String$(18, Chr$(160)) & "" & Trim(RdoAct8!GLACCTNO)
                  !Act04 = String$(18, Chr$(160)) & "" & Trim(RdoAct8!GLDESCR)
                  !Act05 = RdoAct8!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct8!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               iRemCurType = iCurType
               iRemFsLevel = iFsLevel
               If iLevel > 8 Then FillLevel9 Trim(!GLACCTREF)
               If iRemFsLevel = 8 Then
                  With DbAct
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 8
                     !Act02 = "" & Trim(RdoAct8!GLACCTREF)
                     
                     !Act04 = String$(18, Chr$(160)) & "Total " & Trim(RdoAct8!GLDESCR)
                     !Act05 = RdoAct8!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                     .AddNew
                     !Act00 = iRemCurType
                     !Act01 = 8
                     !Act02 = "" & Trim(RdoAct8!GLACCTREF)
                     !Act04 = ""
                     !Act05 = RdoAct8!GLINACTIVE
                     !Act06 = iRemFsLevel
                     .Update
                  End With
               End If
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct8 = Nothing
      Exit Sub
DiaErr1:
      sProcName = "filllevel8"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Public Sub FillLevel9(sMaster As String)
      Dim i As Integer
      Dim RdoAct9 As ADODB.Recordset
      
      On Error GoTo DiaErr1
      sSql = "Qry_FillAccounts '" & sMaster & "'," & Trim(str(iInActive))
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct9)
      If bSqlRows Then
         With RdoAct9
            Do Until .EOF
               With DbAct
                  .AddNew
                  !Act00 = iCurType
                  !Act01 = 9
                  !Act02 = "" & Trim(RdoAct9!GLACCTREF)
                  !Act03 = String$(20, Chr$(160)) & "" & Trim(RdoAct9!GLACCTNO)
                  !Act04 = String$(20, Chr$(160)) & "" & Trim(RdoAct9!GLDESCR)
                  !Act05 = RdoAct9!GLINACTIVE
                  iFsLevel = GetAcctLevel(Trim(RdoAct9!GLACCTREF))
                  !Act06 = iFsLevel
                  .Update
               End With
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      Set RdoAct9 = Nothing
      Exit Sub
DiaErr1:
      sProcName = "filllevel9"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Public Function GetAcctLevel(sAcctRef As String) As Integer
      Dim RdoLvl As ADODB.Recordset
      On Error GoTo DiaErr1
      sSql = "SELECT GLACCTREF,GLFSLEVEL FROM GlacTable " _
             & "WHERE GLACCTREF='" & sAcctRef & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLvl, ES_FORWARD)
      If bSqlRows Then
         GetAcctLevel = RdoLvl!GLFSLEVEL
      Else
         GetAcctLevel = 0
      End If
      Set RdoLvl = Nothing
      Exit Function
      
DiaErr1:
      sProcName = "getacctlevel"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
      
   End Function
   
   Public Sub CreateActTable()
      Dim NewTb1 As TableDef
      Dim NewTb2 As TableDef
      Dim NewTb3 As TableDef
      Dim NewTb4 As TableDef
      Dim NewTb5 As TableDef
      Dim NewTb6 As TableDef
      Dim NewTb7 As TableDef
      Dim NewTb8 As TableDef
      Dim NewFld As Field
      Dim NewIdx1 As Index
      Dim NewIdx2 As Index
      Dim NewIdx3 As Index
      Dim NewIdx4 As Index
      Dim NewIdx5 As Index
      
      
      ' Create fincial statement layout
      On Error Resume Next
      JetDb.Execute "DROP TABLE ActrpTable"
      ' Fields. Note that we allow empties
      Set NewTb1 = JetDb.CreateTableDef("ActrpTable")
      With NewTb1
         'Type
         .Fields.Append .CreateField("Act00", dbInteger)
         'Level
         .Fields.Append .CreateField("Act01", dbInteger)
         'AcctRef
         .Fields.Append .CreateField("Act02", dbText, 12)
         .Fields(2).AllowZeroLength = True
         'Account Number + Spaces to indent
         .Fields.Append .CreateField("Act03", dbText, 32)
         .Fields(3).AllowZeroLength = True
         'Account Desc
         .Fields.Append .CreateField("Act04", dbText, 60)
         .Fields(4).AllowZeroLength = True
         'Active
         .Fields.Append .CreateField("Act05", dbInteger)
         'GLFSLEVEL
         .Fields.Append .CreateField("Act06", dbInteger)
      End With
      
      ' Add the table and indexes to Jet.
      JetDb.TableDefs.Append NewTb1
      Set NewTb2 = JetDb!ActrpTable
      With NewTb2
         Set NewIdx1 = .CreateIndex
         With NewIdx1
            .Name = "AcctTyp"
            .Fields.Append .CreateField("Act00")
         End With
         .Indexes.Append NewIdx1
         
         Set NewIdx2 = .CreateIndex
         With NewIdx2
            .Name = "AcctNo"
            .Fields.Append .CreateField("Act02")
         End With
         .Indexes.Append NewIdx2
      End With
      
      ' Create the current account balance table
      JetDb.Execute "DROP TABLE Current"
      ' Fields. Note that we allow empties
      Set NewTb3 = JetDb.CreateTableDef("Current")
      With NewTb3
         ' AcctRef
         .Fields.Append .CreateField("AcctRef", dbText, 12)
         .Fields(0).AllowZeroLength = True
         ' Current Period
         .Fields.Append .CreateField("AcctBal", dbCurrency)
         ' Current Period %
         .Fields.Append .CreateField("AcctPer", dbCurrency)
      End With
      ' Add the table and indexes to Jet.
      JetDb.TableDefs.Append NewTb3
      Set NewTb4 = JetDb!Current
      With NewTb4
         Set NewIdx3 = .CreateIndex
         With NewIdx3
            .Name = "AcctNo"
            .Fields.Append .CreateField("AcctRef")
         End With
         .Indexes.Append NewIdx3
      End With
      
      ' Create the YTD account balance table
      JetDb.Execute "DROP TABLE YTD"
      ' Fields. Note that we allow empties
      Set NewTb5 = JetDb.CreateTableDef("YTD")
      With NewTb5
         ' AcctRef
         .Fields.Append .CreateField("AcctRef", dbText, 12)
         .Fields(0).AllowZeroLength = True
         ' Current Period
         .Fields.Append .CreateField("AcctBal", dbCurrency)
         ' Current Period %
         .Fields.Append .CreateField("AcctPer", dbCurrency)
      End With
      ' Add the table and indexes to Jet.
      JetDb.TableDefs.Append NewTb5
      Set NewTb6 = JetDb!YTD
      With NewTb6
         Set NewIdx4 = .CreateIndex
         With NewIdx4
            .Name = "AcctNo"
            .Fields.Append .CreateField("AcctRef")
         End With
         .Indexes.Append NewIdx4
      End With
      
      ' Create the Previous account balance table
      JetDb.Execute "DROP TABLE Previous"
      ' Fields. Note that we allow empties
      Set NewTb7 = JetDb.CreateTableDef("Previous")
      With NewTb7
         ' AcctRef
         .Fields.Append .CreateField("AcctRef", dbText, 12)
         .Fields(0).AllowZeroLength = True
         ' Current Period
         .Fields.Append .CreateField("AcctBal", dbCurrency)
         ' Current Period %
         .Fields.Append .CreateField("AcctPer", dbCurrency)
      End With
      ' Add the table and indexes to Jet.
      JetDb.TableDefs.Append NewTb7
      Set NewTb8 = JetDb!Previous
      With NewTb8
         Set NewIdx5 = .CreateIndex
         With NewIdx5
            .Name = "AcctNo"
            .Fields.Append .CreateField("AcctRef")
         End With
         .Indexes.Append NewIdx5
      End With
   End Sub
   
   Private Sub PrintReport()
      Dim sWindows As String
      
      On Error GoTo DiaErr1
      'SetMdiReportsize MdiSect
      sWindows = GetWindowsDir()
      
      MdiSect.crw.DataFiles(0) = sWindows & "\temp\esifina.mdb"
      MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
      MdiSect.crw.Formulas(1) = "Title='Level " & txtLvl _
                           & " Income Statement For Year Beginning " & txtYearBeg & "'"
      MdiSect.crw.Formulas(2) = "Period = 'Period Beginning:  " & txtBeg & " And Ending:  " _
                           & txtEnd & " Ratio Account:  " & lblDsc & "'"
      
      ' To prevent divison by zero
      If cRatio1 = 0 Then cRatio1 = 1
      If cRatio2 = 0 Then cRatio2 = 1
      If cRatio3 = 0 Then cRatio3 = 1
      
      MdiSect.crw.Formulas(3) = "ratio1='" & cRatio1 & "'"
      MdiSect.crw.Formulas(4) = "ratio2='" & cRatio2 & "'"
      MdiSect.crw.Formulas(5) = "ratio3='" & cRatio3 & "'"
      
      MdiSect.crw.ReportFileName = sReportPath & "fingl07.rpt"
      'SetCrystalAction Me
      MouseCursor 0
      Exit Sub
      
DiaErr1:
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors Me
   End Sub
   
   Private Sub optDis_Click()
      BuildAccounts
   End Sub
   
   Private Sub optPrn_Click()
      BuildAccounts
   End Sub
   
   Private Sub txtBeg_DropDown()
      ShowCalendar Me
   End Sub
   
   Private Sub txtend_DropDown()
      ShowCalendar Me
   End Sub
   
   Private Sub txtLvl_LostFocus()
      If Trim(txtLvl) = "" Then txtLvl = 9
   End Sub
   
   Private Sub txtYearBeg_DropDown()
      ShowCalendar Me
   End Sub
