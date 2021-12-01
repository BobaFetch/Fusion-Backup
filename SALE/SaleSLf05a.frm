VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLf05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change A Sales Order Number"
   ClientHeight    =   2790
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Timer tmr1 
      Interval        =   10000
      Left            =   6360
      Top             =   1920
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5940
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Change Sales Order Number"
      Top             =   600
      Width           =   915
   End
   Begin VB.TextBox txtOld 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Current Sales Order"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdDis 
      Caption         =   "&Display"
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Display Sales Order"
      Top             =   2280
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "New Sales Order"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5940
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2790
      FormDesignWidth =   6945
   End
   Begin VB.Label lblOld 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2400
      TabIndex        =   9
      Top             =   1080
      Width           =   252
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Sales Order Number"
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2268
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Sales Order Number"
      Height          =   288
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   2268
   End
   Begin VB.Label lblNew 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   252
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   3720
      TabIndex        =   5
      Top             =   1440
      Width           =   3012
   End
   Begin VB.Label lblCst 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   1212
   End
End
Attribute VB_Name = "SaleSLf05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'1/7/05 Unblocked this function for use with severe restricions
'10/17/05 Fixed Column Not Found (GetSalesOrder)
Option Explicit
Dim bNewExists As Byte
Dim bOldExists As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetLastSalesOrder()
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER FROM SohdTable ORDER BY SONUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet)
   If bSqlRows Then
      With RdoGet
         txtNew = Format$(!SoNumber + 1, SO_NUM_FORMAT)
         ClearResultSet RdoGet
      End With
   Else
      txtNew = SO_NUM_FORMAT
   End If
   Set RdoGet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlastsa"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdChg_Click()
   Dim bByte As Byte
   bByte = CheckSalesOrder()
   If bByte = 1 Then
      MsgBox "Sales Order " & txtOld & " Has Items Attached (Even If Canceled) " & vbCrLf _
         & "And Cannot Be Changed For Integrity Reasons.", _
         vbInformation, Caption
   Else
      bByte = MsgBox("Change The Selected Sales Order Number?", _
              ES_NOQUESTION, Caption)
      If bByte = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         sSql = "UPDATE SohdTable SET SONUMBER=" & Val(txtNew) & " " _
                & "WHERE SONUMBER=" & Val(txtOld) & " "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         If clsADOCon.ADOErrNum = 0 Then
            SysMsg "Sales Order Changed.", True
            Unload Me
         Else
            MsgBox "Could Not Change The Sales Order Number.", _
               vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub cmdDis_Click()
   If bOldExists Then
      PrintReport
   Else
      MsgBox "Requires A Valid Sales Order.", vbInformation, Caption
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2155
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      GetLastSalesOrder
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLf05a = Nothing
   
End Sub


Function GetSalesOrder(lSalesOrder As Variant, bNew As Byte) As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST FROM SohdTable " _
          & "WHERE SONUMBER=" & Trim(str(lSalesOrder)) & " "
   If Not bNew Then sSql = sSql & "AND SOCANCELED=0 "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet)
   If bSqlRows Then
      With RdoGet
         lblOld = "" & Trim(!SOTYPE)
         lblNew = "" & Trim(!SOTYPE)
         If bNew Then txtOld = Format(!SoNumber, SO_NUM_FORMAT)
         lblCst.Alignment = 0
         FindCustomer Me, !SOCUST, False
         ClearResultSet RdoGet
         GetSalesOrder = True
      End With
   Else
      If Not bNew Then
         Beep
         lblCst.Alignment = 1
         lblCst = "****No Such "
         lblNme = "Sales Order Or Sales Order Canceled****"
      End If
      GetSalesOrder = False
   End If
   If bNew And GetSalesOrder Then
      MsgBox "Sales Order Number Is In Use.", vbInformation, Caption
      GetLastSalesOrder
      tmr1.Enabled = True
      If Val(txtNew) > 0 Then GetSalesOrder = True Else GetSalesOrder = False
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getsaleso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub PrintReport()
   Dim lSoNumber As Long
   MouseCursor 13
   On Error GoTo DiaErr1
   lSoNumber = Val(txtOld)
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   sCustomReport = GetCustomReport("sleco1")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   sSql = "{SohdTable.SONUMBER}=" & lSoNumber & " "
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
Private Sub lblcst_Change()
   If Left(lblCst, 6) = "****No" Then
      lblCst.ForeColor = ES_RED
      lblNme.ForeColor = ES_RED
   Else
      lblCst.ForeColor = vbBlack
      lblNme.ForeColor = vbBlack
   End If
   
End Sub

Private Sub tmr1_Timer()
   GetLastSalesOrder
   
End Sub


Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 5)
   txtNew = Format(Abs(Val(txtNew)), SO_NUM_FORMAT)
   tmr1.Enabled = False
   bNewExists = GetSalesOrder(txtNew, True)
   
End Sub


Private Sub txtOld_LostFocus()
   txtOld = CheckLen(txtOld, 5)
   txtOld = Format(Abs(Val(txtOld)), SO_NUM_FORMAT)
   If Val(txtOld) > 0 Then
      bOldExists = GetSalesOrder(txtOld, False)
   Else
      bOldExists = False
      lblCst = ""
      lblNme = ""
   End If
   
End Sub



Private Function CheckSalesOrder() As Byte
   Dim RdoSoi As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT ITSO FROM SoitTable WHERE ITSO=" & Val(txtOld)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoi, ES_FORWARD)
   If bSqlRows Then CheckSalesOrder = 1 Else CheckSalesOrder = 0
   Exit Function
   
DiaErr1:
   sProcName = "checksalesor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
