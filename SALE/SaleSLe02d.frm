VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLe02d 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selling And Credit Info"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox lblCUCOL 
      Height          =   2775
      Left            =   3840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2100
      Width           =   3495
   End
   Begin VB.TextBox lblCUSEL 
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2100
      Width           =   3495
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   0
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   5160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5055
      FormDesignWidth =   7500
   End
   Begin VB.Label lblCreditAvailable 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2580
      TabIndex        =   16
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblCreditLimit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2580
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblCreditExtended 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2580
      TabIndex        =   14
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblPayments 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2580
      TabIndex        =   13
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblOpenSOs 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2580
      TabIndex        =   12
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Extended:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Payments on open SO's:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open SO's:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Available:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2400
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   3600
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3840
      X2              =   7320
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label lblCustomer 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2580
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing/Collection Notes"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Top             =   1740
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Notes:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   1740
      Width           =   1335
   End
End
Attribute VB_Name = "SaleSLe02d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'6/25/04 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdCan_Click()
   Form_Deactivate
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then GetCustomerNotes
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move 500, 1000
   BackColor = ES_ViewBackColor
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SaleSLe02d = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub GetCustomerNotes()
   Dim RdoNte As ADODB.Recordset
   Dim cust As String
   cust = "'" & Compress(lblCustomer) & "'"
   Dim creditLimit As Currency, openSOs As Currency, payments As Currency
   Dim creditExtended As Currency, creditAvailable As Currency
   
   On Error GoTo DiaErr1
   sSql = "SELECT CUREF,CUSEL,CUCOL, CUCREDITLIMIT FROM CustTable" & vbCrLf _
          & "WHERE CUREF=" & cust
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNte, ES_STATIC)
   If bSqlRows Then
      With RdoNte
         lblCUSEL = "" & Trim(!CUSEL)
         lblCUCOL = "" & Trim(!CUCOL)
         creditLimit = !CUCREDITLIMIT
         ClearResultSet RdoNte
      End With
   End If
   
   'get total of open SO's.
   sSql = "select isnull(cast(sum(ITDOLLARS * ITQTY) as decimal(12,0)), 0) as OpenSoTotal" & vbCrLf _
      & "from SohdTable " & vbCrLf _
      & "join SoitTable on ITSO = SONUMBER" & vbCrLf _
      & "left join CihdTable on ITINVOICE = INVNO" & vbCrLf _
      & "where SOCUST = " & cust & vbCrLf _
      & "AND isnull( INVPIF, 0 ) = 0"
   If clsADOCon.GetDataSet(sSql, RdoNte, ES_STATIC) Then
      openSOs = RdoNte!OpenSoTotal
   End If
   
   'get total of payments on open SO's
   sSql = "select isnull(cast(sum(INVPAY) as decimal(12,0)), 0) as PaymentTotal" & vbCrLf _
      & "from CihdTable " & vbCrLf _
      & "where INVCUST = " & cust & vbCrLf _
      & "AND isnull( INVPIF, 1 ) = 0"
   If clsADOCon.GetDataSet(sSql, RdoNte, ES_STATIC) Then
      payments = RdoNte!PaymentTotal
   End If

   Set RdoNte = Nothing
   
   'do arithmetic and display calculations
   creditExtended = openSOs - payments
   creditAvailable = creditLimit - creditExtended
   lblCreditLimit = Format(creditLimit, "###,###,##0")
   lblOpenSOs = Format(openSOs, "###,###,##0")
   lblPayments = Format(payments, "###,###,##0")
   lblCreditExtended = Format(creditExtended, "###,###,##0")
   lblCreditAvailable = Format(creditAvailable, "###,###,##0")
   If creditAvailable < 0 Then
      lblCreditAvailable.ForeColor = ES_RED
   Else
      lblCreditAvailable.ForeColor = Es_TextForeColor
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "getcustomernotes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblCUCOL_Click()
   On Error Resume Next
   cmdCan.SetFocus
   
End Sub

Private Sub lblCUSEL_Click()
   On Error Resume Next
   cmdCan.SetFocus
   
End Sub
