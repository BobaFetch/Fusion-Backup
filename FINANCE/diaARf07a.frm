VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaARf07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update B & O Records"
   ClientHeight    =   2265
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2265
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar prg1 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox txtStart 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "4"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   2280
      Picture         =   "diaARf07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Show Invoices Listed For Export"
      Top             =   1200
      Width           =   350
   End
   Begin VB.CheckBox optVew 
      Caption         =   "vew"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3840
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2265
      FormDesignWidth =   5505
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   3
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
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaARf07a.frx":04DA
      PictureDn       =   "diaARf07a.frx":0620
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   825
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices Found"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label lblFnd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Tag             =   "1"
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "diaARf07a"
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
' diaARf07a - Update B&O Codes For AR Invoice Items
'
' Notes:
'
' Created: 01/29/03 (nth)
' Revisions:
'   06/04/03 (nth) Fixed runtime error per DAP
'
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdUpd_Click()
   GetInvoices
   UpdateBOCodes
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
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   txtstart = Format(Now, "mm/01/yy")
   txtEnd = GetMonthEnd(txtstart)
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaARf07a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
   GetInvoices
End Sub

Private Sub txtstart_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtstart_LostFocus()
   txtstart = CheckDate(txtstart)
End Sub

Private Sub GetInvoices()
   Dim RdoInv As ADODB.Recordset
   sSql = "SELECT COUNT(INVNO) FROM CihdTable WHERE INVDATE >='" _
          & txtstart & "' AND INVDATE <='" & txtEnd & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then
      lblFnd = CStr(RdoInv.Fields(0))
   Else
      lblFnd = "0"
   End If
   Set RdoInv = Nothing
End Sub

Private Sub UpdateBOCodes()
   Dim RdoInv As ADODB.Recordset
   Dim rdoCod As ADODB.Recordset
   Dim I As Integer
   
   On Error GoTo DiaErr1
   sSql = "SELECT ITBOSTATIC,ITBOSTATE,ITBOCODE,ITPART " _
          & "FROM SoitTable INNER JOIN CihdTable " _
          & "ON SoitTable.ITINVOICE = CihdTable.INVNO " _
          & "WHERE INVDATE >='" & txtstart & "' AND INVDATE <='" _
          & txtEnd & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC)
   If bSqlRows Then
      With RdoInv
         prg1.max = .RecordCount + 1
         prg1.Visible = True
         While Not .EOF
            'Retail
            sSql = "SELECT TAXRATE,TAXCODE,TAXSTATE FROM PartTable " _
                   & "INNER JOIN TxcdTable ON PartTable.PABORTAX = " _
                   & "TxcdTable.TAXREF WHERE PARTREF = '" & !ITPART & "'"
            bSqlRows = clsADOCon.GetDataSet(sSql, rdoCod, ES_FORWARD)
            If bSqlRows Then
               !ITBOSTATIC = rdoCod!TAXRATE
               !ITBOSTATE = Trim(rdoCod!taxState)
               !ITBOCODE = Trim(rdoCod!taxCode)
               .Update
               Set rdoCod = Nothing
            Else
               'Wholesale
               sSql = "SELECT TAXRATE,TAXCODE,TAXSTATE FROM PartTable " _
                      & "INNER JOIN TxcdTable ON PartTable.PABOWTAX = " _
                      & "TxcdTable.TAXREF WHERE PARTREF = '" & !ITPART & "'"
               bSqlRows = clsADOCon.GetDataSet(sSql, rdoCod, ES_FORWARD)
               If bSqlRows Then
                  !ITBOSTATIC = rdoCod!TAXRATE
                  !ITBOSTATE = Trim(rdoCod!taxState)
                  !ITBOCODE = Trim(rdoCod!taxCode)
                  .Update
                  Set rdoCod = Nothing
               End If
            End If
            .MoveNext
            I = I + 1
            prg1.Value = I
            DoEvents
         Wend
         prg1.Visible = False
      End With
   End If
   Set RdoInv = Nothing
   sMsg = "Successfully Updated B & 0 Codes"
   MsgBox sMsg, vbInformation, Caption
   Exit Sub
DiaErr1:
   sProcName = "updatebocodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
