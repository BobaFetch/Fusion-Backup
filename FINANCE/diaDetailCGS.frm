VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaDetailCGS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detail Cost Of Goods Sold (Report)"
   ClientHeight    =   3630
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optSORet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   2640
      Width           =   855
   End
   Begin VB.CheckBox optSum 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox optSO 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   11
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
      PictureUp       =   "diaDetailCGS.frx":0000
      PictureDn       =   "diaDetailCGS.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   12
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
      PictureUp       =   "diaDetailCGS.frx":028C
      PictureDn       =   "diaDetailCGS.frx":03D2
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show SO returns"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   21
      Top             =   2640
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Journal Only"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Numbers"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   2160
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descriptions"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaDetailCGS"
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

'*********************************************************************************
' diaDetailCGS - Cost Of Goods Sold
'
' Notes:
'
' Created: 02/18/04 (nth)
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
 '  txtEnd = Format(GetServerDateTime(), "mm/dd/yy")
 '  txtBeg = Format(txtEnd, "mm/01/yy")
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
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
   Set diaDetailCGS = Nothing
End Sub

Private Function DataReady() As Boolean
      
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Dim RdoMODetail As ADODB.Recordset
   Dim strStDate As String
   Dim strEndDate As String
   strStDate = txtBeg
   strEndDate = txtEnd
   
   DataReady = False
   sSql = "TRUNCATE TABLE EsMOPartsCostDetail"
   'RdoCon.Execute sSql, rdExecDirect
   clsADOCon.ExecuteSQL sSql
     
   'sSql = "RptPSMOs '" & strStDate & "', '" & strEndDate & "'"
   'RdoCon.Execute sSql, rdExecDirect
   'clsADOCon.ExecuteSQL sSql

   sSql = "SELECT DISTINCT INVNO, SoitTable.ITPSNUMBER as ITPSNUMBER, SoitTable.ITPSITEM as ITPSITEM " & vbCrLf
   sSql = sSql & "      From " & vbCrLf
   sSql = sSql & "         (((((CihdTable CihdTable INNER JOIN SoitTable SoitTable ON" & vbCrLf
   sSql = sSql & "            CihdTable.INVNO = SoitTable.ITINVOICE)" & vbCrLf
   sSql = sSql & "          INNER JOIN PartTable PartTable ON" & vbCrLf
   sSql = sSql & "            SoitTable.ITPART = PartTable.PARTREF)" & vbCrLf
   sSql = sSql & "         INNER JOIN SohdTable SohdTable ON" & vbCrLf
   sSql = sSql & "         SoitTable.ITSO = SohdTable.SONUMBER)" & vbCrLf
   sSql = sSql & "         INNER JOIN InvaTable InvaTable ON" & vbCrLf
   sSql = sSql & "            SoitTable.ITSO = InvaTable.INSONUMBER AND" & vbCrLf
   sSql = sSql & "         SoitTable.ITNUMBER = InvaTable.INSOITEM AND" & vbCrLf
   sSql = sSql & "         SoitTable.ITREV = InvaTable.INSOREV)" & vbCrLf
   sSql = sSql & "          LEFT OUTER JOIN PsitTable PsitTable ON" & vbCrLf
   sSql = sSql & "            InvaTable.INPSNUMBER = PsitTable.PIPACKSLIP AND" & vbCrLf
   sSql = sSql & "         InvaTable.INPSITEM = PsitTable.PIITNO AND" & vbCrLf
   sSql = sSql & "         InvaTable.INPART = PsitTable.PIPART AND" & vbCrLf
   sSql = sSql & "         InvaTable.INSONUMBER = PsitTable.PISONUMBER AND" & vbCrLf
   sSql = sSql & "         InvaTable.INSOITEM = PsitTable.PISOITEM AND" & vbCrLf
   sSql = sSql & "         InvaTable.INSOREV = PsitTable.PISOREV)" & vbCrLf
   sSql = sSql & "          INNER JOIN PshdTable PshdTable ON" & vbCrLf
   sSql = sSql & "            PsitTable.PIPACKSLIP = PshdTable.PsNumber" & vbCrLf
   sSql = sSql & "      Where" & vbCrLf
   sSql = sSql & "         CihdTable.INVDATE Between '" & strStDate & "' and '" & strEndDate & "' AND" & vbCrLf
   sSql = sSql & "--       INPSNUMBER = 'PS020139' AND" & vbCrLf
   sSql = sSql & "         (InvaTable.INTYPE = 4 OR" & vbCrLf
   sSql = sSql & "         InvaTable.INTYPE = 3 OR" & vbCrLf
   sSql = sSql & "         InvaTable.INTYPE = 26 OR" & vbCrLf
   sSql = sSql & "         InvaTable.INTYPE = 25 OR" & vbCrLf
   sSql = sSql & "         InvaTable.INTYPE = 24) AND" & vbCrLf
   sSql = sSql & "         PshdTable.PSCANCELED = 0 AND" & vbCrLf
   sSql = sSql & "         CihdTable.INVCANCELED = 0 AND" & vbCrLf
   sSql = sSql & "         SoitTable.ITCANCELED = 0"

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMODetail)
   If bSqlRows Then
      With RdoMODetail
         Do Until .EOF
            
            Dim InvNo As Long
            Dim strPSNum As String
            Dim PSItem As Integer
            
            InvNo = Trim(!InvNo)
            strPSNum = "" & Trim(!ITPSNUMBER)
            PSItem = Trim(!ITPSITEM)
            
            sSql = "RptPSMOs '" & CStr(InvNo) & "', '" & strPSNum & "','" & CStr(PSItem) & "'"
            clsADOCon.ExecuteSQL sSql ', rdExecDirect
            
'            sSql = "RptPSMOs '" & strStDate & "', '" & strEndDate & "'"
'            RdoCon.Execute sSql, rdExecDirect
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   
   Set RdoMODetail = Nothing

   If (Err.Number = 0) Then DataReady = True
   Exit Function
DiaErr1:
   sProcName = "DataReady"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optDis_Click()
   If DataReady() Then
      PrintReport
   End If
End Sub

Private Sub optPrn_Click()
   If DataReady() Then
      PrintReport
   End If
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim b As Byte
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   sCustomReport = GetCustomReport("fincgsCostDetail.rpt")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Dsc"
   aFormulaName.Add "Ext"
   aFormulaName.Add "SO"
   aFormulaName.Add "SORet"
    
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'From " & CStr(txtBeg & " Through " & txtEnd) & "'")
   aFormulaValue.Add optDsc
   aFormulaValue.Add optExt
   aFormulaValue.Add optSO
   aFormulaValue.Add optSORet
   
   sSql = "{CihdTable.INVDATE} >= cdate('" & txtBeg _
          & "') and {CihdTable.INVDATE} <= cdate('" & txtEnd & "')"
        sSql = sSql & " and {InvaTable.INTYPE} in [24.00, 25.00, 26.00, 3.00, 4.00] and " _
                    & "{PshdTable.PSCANCELED} = 0 and " _
                    & "{CihdTable.INVCANCELED} = 0 and " _
                    & "{SoitTable.ITPSNUMBER} = {InvaTable.INPSNUMBER} and " _
                    & "{SoitTable.ITCANCELED} = 0.00"
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtBeg.Text) & Trim(txtEnd.Text) & Trim(optDsc.Value) & Trim(optExt.Value) _
              & Trim(optSO.Value) & Trim(optSum.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   On Error Resume Next
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   
   If Len(Trim(sOptions)) > 0 Then
     If dToday < 21 Then
      txtBeg = Mid(sOptions, 1, 8)
      txtEnd = Mid(sOptions, 9, 8)
     Else
      txtBeg = Format(Now, "mm/01/yy")
      txtEnd = GetMonthEnd(txtBeg)
     End If
      optDsc.Value = Mid(sOptions, 17, 1)
      optExt.Value = Mid(sOptions, 18, 1)
      optSO.Value = Mid(sOptions, 19, 1)
      optSum.Value = Mid(sOptions, 20, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub
