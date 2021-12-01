VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form diaSCf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Proposed Standard Cost For All Parts"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5040
      Top             =   240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4170
      FormDesignWidth =   7215
   End
   Begin VB.CheckBox optPrn 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox chkStd 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   2040
      Width           =   855
   End
   Begin VB.CheckBox chkB 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.CheckBox chkLab 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.CheckBox chkExp 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Tag             =   "1"
      Top             =   2400
      Width           =   555
   End
   Begin ComctlLib.ProgressBar prg1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      ToolTipText     =   "Update Standard Cost To Calculated Total"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaSCf01a.frx":0000
      PictureDn       =   "diaSCf01a.frx":0146
   End
   Begin Threed.SSRibbon SSRibbon1 
      Height          =   255
      Left            =   0
      TabIndex        =   18
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
      PictureUp       =   "diaSCf01a.frx":028C
      PictureDn       =   "diaSCf01a.frx":03D2
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   19
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
      PictureUp       =   "diaSCf01a.frx":0524
      PictureDn       =   "diaSCf01a.frx":066A
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblRec 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      Height          =   285
      Index           =   12
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Of"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Else report will display in preview)"
      Height          =   285
      Index           =   10
      Left            =   4200
      TabIndex        =   23
      Top             =   2760
      Width           =   2625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Treated Like Raw Material Otherwise)"
      Height          =   285
      Index           =   9
      Left            =   4200
      TabIndex        =   22
      Top             =   1680
      Width           =   2865
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Primary & Secondary Shops Otherwise)"
      Height          =   285
      Index           =   8
      Left            =   4200
      TabIndex        =   21
      Top             =   1320
      Width           =   2865
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Standard Exp Used Otherwise)"
      Height          =   285
      Index           =   7
      Left            =   4200
      TabIndex        =   17
      Top             =   960
      Width           =   2625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Print This Update"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   2865
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percent (Zero To Ignore)"
      Height          =   285
      Index           =   5
      Left            =   4200
      TabIndex        =   10
      Top             =   2400
      Width           =   2385
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Variances Greater Than Or Equal To"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2985
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Standard Equal to The Proposed?"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2985
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Based ON BOM for ""B"" Parts?"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2865
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Labor Cost From The Routings"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2865
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Expense Cost From Routings"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2505
   End
End
Attribute VB_Name = "diaSCf01a"
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

'**************************************************************************************
' diaSCf01a - Standard cost all parts
'
' Notes: Requires diaSCp02a to run.
'
' Created: 12/11/01 (nth)
' Revisions:
'   12/04/02 (nth) Add record count and progress bar
'   12/05/02 (nth) Added variance report
'
'**************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'**************************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdUpd_Click()
   UpdateAllParts
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
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
   On Error Resume Next
   FormUnload
   Set diaSCf01a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub UpdateAllParts()
   Dim bSuccess As Byte
   Dim RdoAllLevels As ADODB.Recordset
   Dim rdoCnt As ADODB.Recordset
   Dim RdoPrt As ADODB.Recordset
   Dim lCount As Long
   Dim lRec As Long
   Dim sPart As String
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   ' Get number of parts to cost
   sSql = "SELECT COUNT(PARTREF) FROM PartTable WHERE " _
          & "(PALEVEL IN ('1','2','3','4','7'))"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCnt, ES_FORWARD)
   If bSqlRows Then
      With rdoCnt
         lCount = .Fields(0)
      End With
   End If
   Set rdoCnt = Nothing
   
   lblCount = lCount
   lblRec = 0
   prg1.max = lCount
   
   prg1.Visible = True
   lblRec.Visible = True
   lblCount.Visible = True
   z1(12).Visible = True
   z1(11).Visible = True
   
   DoEvents
   
   ' Get list of parts to cost
   sSql = "SELECT PARTREF FROM PartTable WHERE " _
          & "(PALEVEL IN ('1','2','3','4','7'))"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   
   If bSqlRows Then
      
      While Not RdoPrt.EOF
         
         lRec = lRec + 1
         prg1.Value = lRec
         lblRec = lRec
         DoEvents
         
         sPart = RdoPrt.Fields(0)
         diaSCp02a.GetBillOfMaterial sPart, ""
         
         If chkStd = vbChecked Then
            
            sSql = "SELECT DISTINCT BomPartRef,BomLevel FROM EsReportTmpBomTable " _
             & "WHERE BomPartRef IS NOT NULL ORDER BY BomLevel DESC"
      
            bSqlRows = clsADOCon.GetDataSet(sSql,RdoAllLevels, ES_FORWARD)
            If bSqlRows Then
               With RdoAllLevels
                  bSuccess = 1
                  While Not .EOF And bSuccess = 1
                     bSuccess = diaSCp02a.StdCostPart(Trim(!BomPartRef))
                     .MoveNext
                  Wend
               End With
            End If
            Set RdoAllLevels = Nothing
            
            ' Now cost the actual assembly
            If bSuccess Then
               bSuccess = diaSCp02a.StdCostPart(sPart)
            End If
            
            If bSuccess <> 1 Then
               sMsg = "Error in costing " & sPart & vbCrLf _
                      & "Standard Cost Rollup Canceled"
               MsgBox sMsg, vbExclamation
               Set RdoAllLevels = Nothing
               Set RdoPrt = Nothing
               lblRec.Visible = False
               lblCount.Visible = False
               z1(11).Visible = False
               z1(12).Visible = False
               prg1.Visible = False
               Exit Sub
               
            End If
         End If
         RdoPrt.MoveNext
      Wend
      
      sMsg = "Standard Cost Updated For All Parts"
      SysMsg sMsg, True
   End If
   
   
   Set RdoPrt = Nothing
   
   lblRec.Visible = False
   lblCount.Visible = False
   z1(11).Visible = False
   z1(12).Visible = False
   prg1.Visible = False
   
   If optPrn Then
      Dim sCustomReport As String
      Dim cCRViewer As EsCrystalRptViewer
      Dim aFormulaValue As New Collection
      Dim aFormulaName As New Collection
      Dim strSumDetail As String
      
      sCustomReport = GetCustomReport("finsc02f")
      
      Set cCRViewer = New EsCrystalRptViewer
      cCRViewer.Init
      cCRViewer.SetReportFileName sCustomReport, sReportPath
      
      cCRViewer.SetReportTitle = "finsc02f.rpt"
      cCRViewer.ShowGroupTree False
      
      aFormulaName.Add "CompanyName"
      aFormulaName.Add "RequestBy"
      aFormulaName.Add "VARIANCE"
      aFormulaName.Add "UPDSTD"
      aFormulaName.Add "Title1"
      
      aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
      aFormulaValue.Add CStr("'" & sInitials & "'")
      aFormulaValue.Add CStr("'" & txtPer & "'")
      
      ' Title 1
      If chkStd Then
         aFormulaValue.Add CStr("'1'")
         aFormulaValue.Add CStr("''")
      Else
         aFormulaValue.Add CStr("'0'")
         aFormulaValue.Add CStr("'Update Of Standard Costs Requested'")
      End If
      
      ' Title 2
      If Val(txtPer) <> 0 Then
         aFormulaName.Add "Title2"
         aFormulaValue.Add CStr("'List Proposed Costs With A Variance Greater Than Or Equal To " _
                              & Val(txtPer) & " Percent'")
      End If
      
      aFormulaName.Add "Title3"
      
      sMsg = "Lab From Routings? "
      If chkLab Then sMsg = sMsg & "Y " Else sMsg = sMsg & "N "
      sMsg = sMsg & "Exp From Routings? "
      If chkExp Then sMsg = sMsg & "Y " Else sMsg = sMsg & "N "
      sMsg = sMsg & "Update From BOM For B Parts? "
      If chkB Then sMsg = sMsg & "Y " Else sMsg = sMsg & "N "
      ' Title 3
      aFormulaValue.Add CStr("'" & sMsg & "'")
      
      cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
      
      cCRViewer.CRViewerSize Me
      ' Set report parameter
      cCRViewer.SetDbTableConnection True
      'cCRViewer.SetTableConnection aRptPara
      cCRViewer.OpenCrystalReportObject Me, aFormulaName
      
      cCRViewer.ClearFieldCollection aFormulaName
      cCRViewer.ClearFieldCollection aFormulaValue
   
   End If
   
   MouseCursor 0
   Exit Sub
   
   ' Error handeling
DiaErr1:
   Set RdoAllLevels = Nothing
   Set RdoPrt = Nothing
   sProcName = "UpdateAllParts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
Public Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(chkExp.Value) _
              & RTrim(chkLab.Value) _
              & RTrim(chkB.Value) _
              & RTrim(chkStd.Value) _
              & RTrim(optPrn.Value) _
              & RTrim(txtPer)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      chkExp.Value = Val(Left(sOptions, 1))
      chkLab.Value = Val(Mid(sOptions, 2, 1))
      chkB.Value = Val(Mid(sOptions, 3, 1))
      chkStd.Value = Val(Mid(sOptions, 4, 1))
      optPrn.Value = Val(Mid(sOptions, 5, 1))
      txtPer = Mid(sOptions, 6, Len(sOptions) - 5)
   Else
      chkExp.Value = vbUnchecked
      chkLab.Value = vbUnchecked
      chkB.Value = vbUnchecked
      chkStd.Value = vbUnchecked
      optPrn.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

