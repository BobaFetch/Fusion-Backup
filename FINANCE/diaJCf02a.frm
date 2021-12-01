VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaJCf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Manufacturing Order Labor Hours"
   ClientHeight    =   3780
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Tag             =   "4"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Contains Qualified Runs"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Qualified Part Numbers (CO)"
      Top             =   720
      Width           =   3060
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1065
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   9
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
      PictureUp       =   "diaJCf02a.frx":0000
      PictureDn       =   "diaJCf02a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3780
      FormDesignWidth =   7260
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   6600
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Status"
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   19
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   17
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   15
      Top             =   2205
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   2205
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   13
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblStu 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   12
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Top             =   765
      Width           =   975
   End
End
Attribute VB_Name = "diaJCf02a"
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
'
' diaJCf02a - MO Cost Analysis Report
'
' Created: (cjs)
' Revisions:
'   06/11/03 (nth) Added VITADDRES to PO cost on report per incident 17887
'   05/07/04 (nth) Removed jet DB logic use subreport instead
'   05/18/04 (nth) Added options from MCS see dbm23
' 4/11/05 TEL - formatted date passed to MO Cost Analysis (finjc01.rpt) as mm/dd/yy
' 6/8/05 TEL - allow selection of closed runs
'
'************************************************************************************

'Dim RdoQry As rdoQuery
Dim AdoCmdObj As ADODB.Command
Dim bOnLoad As Byte
Dim bGoodMo As Byte

Dim lRunno As Long
Dim SPartRef As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

Public Sub GetStatus()
   Dim RdoStu As ADODB.Recordset
   On Error GoTo DiaErr1
   SPartRef = Compress(cmbPrt)
   sSql = "SELECT RUNSTATUS from RunsTable WHERE RUNREF = '" _
          & SPartRef & "' AND RUNNO=" & cmbRun & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoStu, ES_FORWARD)
   If bSqlRows Then
      With RdoStu
         lblStu = "" & Trim(!RUNSTATUS)
         .Close
      End With
   Else
      lblStu = ""
   End If
   Set RdoStu = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getstatus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillFormRuns()
   Dim RdoRns As ADODB.Recordset
   Dim SPartRef As String
   cmbRun.Clear
   SPartRef = Compress(cmbPrt)
   'RdoQry(0) = SPartRef
   AdoCmdObj.parameters(0) = SPartRef
   bSqlRows = clsADOCon.GetQuerySet(RdoRns, AdoCmdObj)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr cmbRun.hWnd, Format$(!Runno, "####0")
            .MoveNext
         Loop
      End With
   Else
   End If
   If cmbRun.ListCount > 0 Then
      cmbRun = Format(cmbRun.List(0), "####0")
      GetStatus
   End If
   On Error Resume Next
   Set RdoRns = Nothing
   
   Exit Sub
   
DiaErr1:
   sProcName = "fillformru"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbPrt_Click()
   LocalFindPart Me
   FillFormRuns
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Trim(cmbPrt) = "" Then
      cmbPrt = "ALL"
      lblDsc = "*** Part Number Wasn't Found ***"
   Else
    lblDsc = ""
    LocalFindPart Me
    FillFormRuns
   End If
   
End Sub


Private Sub cmbRun_Click()
    If Val(cmbRun) > 0 Then GetStatus Else _
       lblStu = ""
   
End Sub


Private Sub cmbRun_LostFocus()

   If Trim(cmbRun) = "" Then
      cmbRun = "ALL"
   Else
    If Val(cmbRun) > 0 Then GetStatus Else _
           lblStu = ""
  End If
  
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombo()
   Dim RdoPcl As ADODB.Recordset
   Dim sTempPart As String
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,PADESC,PALEVEL,RUNREF," _
          & "RUNSTATUS FROM PartTable,RunsTable WHERE " _
          & "RUNREF=PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPcl)
   If bSqlRows Then
      With RdoPcl
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         Do Until .EOF
            If sTempPart <> Trim(!PARTNUM) Then
               'cmbPrt.AddItem "" & Trim(!PARTNUM)
               AddComboStr cmbPrt.hWnd, "" & Trim(!PARTNUM)
               sTempPart = Trim(!PARTNUM)
            End If
            .MoveNext
         Loop
      End With
      If cmbPrt.ListCount > 0 Then FillFormRuns
   Else
      MsgBox "No Matching Runs Recorded.", _
         vbInformation, Caption
   End If
   On Error Resume Next
   Set RdoPcl = Nothing
   cmbPrt.SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdSearch_Click()
   fileDlg.Filter = "Excel File (*.xls) | *.xls"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.filename
   End If

End Sub

Private Sub cmbExport_Click()
    If (txtFilePath <> "") Then
        ExportTime
    Else
        MsgBox "Please Select The FileName", vbOKOnly
    End If
        
End Sub

Private Function ExportTime()

   Dim sParts As String
   Dim sBegDate As String
   Dim sEnddate As String
   Dim sFileName As String
   Dim sRun As String
   
   On Error GoTo ExportError

   Dim rdoPo As ADODB.Recordset
   Dim i As Integer
   Dim sFieldsToExport(11) As String
   
   AddFieldsToExport sFieldsToExport
   
    If Trim(txtBeg) = "" Then txtBeg = "ALL"
    If Trim(txtEnd) = "" Then txtEnd = "ALL"
    If Not IsDate(txtBeg) Then
       sBegDate = "01/01/2000"
    Else
       sBegDate = Format(txtBeg, "mm/dd/yyyy")
    End If
    If Not IsDate(txtEnd) Then
       sEnddate = "12/31/2024"
    Else
       sEnddate = Format(txtEnd, "mm/dd/yyyy")
    End If

    
    If (Trim(cmbPrt) = "" Or cmbPrt = "ALL") Then sParts = "%" Else sParts = Compress(cmbPrt)
    If (Trim(cmbRun) = "" Or cmbRun = "ALL") Then sRun = "%" Else sRun = cmbRun
    


    sSql = "select TCPARTREF, TCRUNNO,TCHOURS, TCCODE, TCSHOP,TCWC, " & vbCrLf
    sSql = sSql & " TCOHFIXED,TMDAY,PREMNUMBER, PREMLSTNAME, PREMFSTNAME " & vbCrLf
    sSql = sSql & " FROM (TcitTable TcitTable INNER JOIN EmplTable EmplTable ON " & vbCrLf
    sSql = sSql & "        TCEMP = EmplTable.PREMNUMBER) " & vbCrLf
    sSql = sSql & "     INNER JOIN TchdTable TchdTable ON " & vbCrLf
    sSql = sSql & "        TCCARD = TchdTable.TMCARD " & vbCrLf
    sSql = sSql & " WHERE TCPARTREF LIKE '" & sParts & "' AND TCRUNNO LIKE '" & sRun & "'" & vbCrLf
    sSql = sSql & " AND TchdTable.TMDAY between '" & sBegDate & "' AND '" & sEnddate & "' " & vbCrLf
    sSql = sSql & " AND  TCPARTREF <> '' " & vbCrLf
    sSql = sSql & " ORDER BY TCRUNNO ASC "

   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPo, ES_STATIC)
   
   If bSqlRows Then
      sFileName = txtFilePath.Text
      SaveAsExcel rdoPo, sFieldsToExport, sFileName
   Else
      MsgBox "No records found. Please try again.", vbOKOnly
   End If

   Set rdoPo = Nothing
   Exit Function
   
ExportError:
   MouseCursor 0
   cmbExport.enabled = True
   MsgBox Err.Description
   

End Function

Private Function AddFieldsToExport(ByRef sFieldsToExport() As String)
   
   Dim i As Integer
   i = 0
   sFieldsToExport(i) = "TCPARTREF"
   sFieldsToExport(i + 1) = "TCRUNNO"
   sFieldsToExport(i + 2) = "TCHOURS"
   sFieldsToExport(i + 3) = "TCCODE"
   sFieldsToExport(i + 4) = "TCSHOP"
   sFieldsToExport(i + 5) = "TCWC"
   sFieldsToExport(i + 6) = "TCOHFIXED"
   sFieldsToExport(i + 7) = "TMDAY"
   sFieldsToExport(i + 8) = "PREMNUMBER"
   sFieldsToExport(i + 9) = "PREMLSTNAME"
   sFieldsToExport(i + 10) = "PREMFSTNAME"
End Function


Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim i As Integer
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   '    sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
   '        & "RunsTable WHERE RUNREF = ? " _
   '        & "AND (RUNSTATUS<>'CA' AND RUNSTATUS<>'CL')  "
   
   'allow selection of closed runs
   sSql = "SELECT RUNREF,RUNSTATUS,RUNNO FROM " _
          & "RunsTable WHERE RUNREF = ? " _
          & "AND (RUNSTATUS<>'CA')  "
   
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmRunRef As ADODB.Parameter
   Set prmRunRef = New ADODB.Parameter
   prmRunRef.Type = adChar
   prmRunRef.SIZE = 30
   AdoCmdObj.parameters.Append prmRunRef
   
   
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
'   txtBeg = ""
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoCmdObj = Nothing
   Set diaJCf02a = Nothing
End Sub
Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub SaveOptions()
   Dim sOptions As String

End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next

End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
      lblStu = ""
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optAct_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
End Sub


Private Sub LocalFindPart(frm As Form, Optional sGetPart As String)
   Dim RdoPrt As ADODB.Recordset
   If sGetPart = "" Then
      sGetPart = Compress(frm.cmbPrt)
   Else
      sGetPart = Compress(sGetPart)
   End If
   On Error Resume Next
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
             & "WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
      If bSqlRows Then
         With RdoPrt
            frm.cmbPrt = "" & Trim(!PARTNUM)
            frm.lblDsc.ForeColor = frm.ForeColor
            frm.lblDsc = "" & Trim(!PADESC)
         End With
      Else
         frm.lblDsc.ForeColor = ES_RED
         frm.cmbPrt = "NONE"
         frm.lblDsc = "*** Part Number Wasn't Found ***"
         
      End If
   Else
      frm.cmbPrt = "NONE"
   End If
   Set RdoPrt = Nothing
End Sub


