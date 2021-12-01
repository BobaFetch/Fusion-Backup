VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ShopSHF10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SINC Reports - Deliveries from Supply Base"
   ClientHeight    =   1770
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbCreate 
      Caption         =   "Create Report"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHF10a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Tag             =   "4"
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   1425
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5160
      Top             =   1440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1770
      FormDesignWidth =   6045
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   4320
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "ShopSHF10a"
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
Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "shf09", sOptions)
   If Trim(txtEnd) = "" Then txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   If Trim(txtBeg) = "" Then txtBeg = Format(DateAdd("d", -6, ES_SYSDATE), "mm/dd/yy")
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sBeg As String * 8
   Dim sEnd As String * 8
   
   sBeg = txtBeg
   sEnd = txtEnd
   SaveSetting "Esi2000", "EsiProd", "shf09", Trim(sOptions)
   
End Sub

Private Sub cmbCreate_Click()

   Dim strExFileName As String
   
   ' Clear the data
   fileDlg.filename = ""
   fileDlg.Filter = "comma delimited file (*.csv) | *.csv"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
      strExFileName = ""
      MsgBox "Please select report file name.", vbOKOnly
      Exit Sub
   Else
       strExFileName = fileDlg.filename
   End If
   
   MouseCursor ccHourglass
   
   CreateSINCDelSupplyBase strExFileName
   MouseCursor ccArrow
   Exit Sub


End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopSHF10a = Nothing
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDate(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDate(txtEnd)
   
End Sub

Public Function CreateSINCDelSupplyBase(strFileName As String)
'   Dim RdoSinc As rdoResultset
   Dim RdoSinc As ADODB.Recordset
   On Error GoTo modErr1
   
   
   If (txtBeg = "" Or txtEnd = "") Then
      MsgBox "Please select Starting and Ending Dates.", vbOKOnly
      Exit Function
   End If
   
   Dim nFileNum As Integer
   
   ' Get a free file number
   nFileNum = FreeFile
   Open strFileName For Output As nFileNum
   
   ' Add header
   Dim strHeader  As String
   Dim strLine As String
   Dim strBestCode As String
   Dim strUnk As String
   Dim strSEA As String
   Dim strRptDate As String
   Dim strCus As String
   Dim strPartNum As String
   Dim cOntimeQty As Double
   Dim cLateQty As Double
   Dim cTotQty As Double
   Dim cTotLateQty As Double
   Dim cGrdTotalQty As Double
   
   
   strUnk = "UNKNOWN"
   strSEA = "SEA"
   strBestCode = "BE10048503"
   cTotLateQty = 0
   cGrdTotalQty = 0
   
   strHeader = "ReportDate,SupplierBESTCode,BoeingBusinessUnit,BoeingSite,BoeingProgram,BoeingPartNumber,SupplierPartNumber,UnitOfMeasure,QuantityLateSubTierDeliveries,TotalQuantitySubTierDeliveries"
   
   Print #nFileNum, strHeader
   
   strRptDate = Format(txtEnd, "yyyy-mm-dd")

   sSql = "Sinc_DeliverySupplyBase '" & txtBeg & "','" & txtEnd & "'"
'   bSqlRows = GetDataSet(RdoSinc, ES_KEYSET)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSinc, ES_KEYSET)
   
   If bSqlRows Then
      On Error Resume Next
      With RdoSinc
         Do Until .EOF
            
            '2009-09-25,BE19037363,BCA,SEA,777,RJX986432-300,,LT,1,16
            '2009-09-25,BE19037363,site,site,site,site,,,5,153
            
            strCus = IIf(IsNull(.Fields(1)), strUnk, .Fields(1))
            strPartNum = "" & Trim(.Fields(0))
            cOntimeQty = IIf(IsNull(.Fields(2)), 0, CDbl(.Fields(2)))
            cLateQty = IIf(IsNull(.Fields(3)), 0, CDbl(.Fields(3)))
            cTotQty = IIf(IsNull(.Fields(4)), 0, CDbl(.Fields(4)))
         
            strLine = strRptDate & "," & strBestCode & "," & strCus _
                     & "," & strSEA & "," & strUnk & "," & strPartNum _
                     & ",,," & CStr(cLateQty) & "," & CStr(cTotQty)
            
            cTotLateQty = cTotLateQty + cLateQty
            cGrdTotalQty = cGrdTotalQty + cTotQty
                     
            Print #nFileNum, strLine
            
            .MoveNext
         Loop
         ClearResultSet RdoSinc
      End With
      
      ' Enter the last entry
      strLine = strRptDate & "," & strBestCode & ",site,site,site,site,,," & CStr(cTotLateQty) & "," & CStr(cGrdTotalQty)
      Print #nFileNum, strLine
      
   End If
   
   
   Set RdoSinc = Nothing
   
   ' Close the filename
   Close nFileNum
   
   Exit Function
   
modErr1:
   sProcName = "CreateSINCDelSupplyBase"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function


