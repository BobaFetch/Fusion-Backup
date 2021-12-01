VERSION 5.00
Begin VB.Form ShopSHe02i 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Link to a higher level MO"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboHigherMoPart 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   840
      Width           =   3545
   End
   Begin VB.ComboBox cboHigherMoRun 
      Height          =   315
      Left            =   5340
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrePick 
      Caption         =   "Link"
      Default         =   -1  'True
      Height          =   435
      Left            =   2160
      TabIndex        =   2
      Top             =   2580
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   435
      Left            =   3840
      TabIndex        =   3
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Label Label7 
      Caption         =   "Type"
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblLowerType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   20
      Top             =   420
      Width           =   375
   End
   Begin VB.Label lblHigherQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5340
      TabIndex        =   19
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Qty"
      Height          =   255
      Left            =   4860
      TabIndex        =   18
      Top             =   1260
      Width           =   315
   End
   Begin VB.Label lblHigherType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   17
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label lblLowerMoDescription 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1260
      TabIndex        =   15
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Link MO to"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   900
      Width           =   975
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   4860
      TabIndex        =   13
      Top             =   900
      Width           =   435
   End
   Begin VB.Label lblHigherStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblHigherDescription 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1260
      TabIndex        =   11
      Top             =   1260
      Width           =   3495
   End
   Begin VB.Label lblLowerMoStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Qty"
      Height          =   255
      Left            =   4860
      TabIndex        =   9
      Top             =   480
      Width           =   315
   End
   Begin VB.Label lblLowerMoQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5340
      TabIndex        =   8
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Run"
      Height          =   255
      Left            =   4860
      TabIndex        =   7
      Top             =   180
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Lower MO"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   915
   End
   Begin VB.Label lblLowerMoRun 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5340
      TabIndex        =   5
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lblLowerMoPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1260
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "ShopSHe02i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bLoading As Boolean
'Private whereClause As String

Private Sub cboHigherMoPart_Change()
   FillMoPartInfo cboHigherMoPart, lblHigherDescription, lblHigherType
   FillMoRunCombo cboHigherMoPart, Me.cboHigherMoRun, GetWhereClause
End Sub

Private Sub cboHigherMoPart_Click()
   FillMoPartInfo cboHigherMoPart, lblHigherDescription, lblHigherType
   FillMoRunCombo cboHigherMoPart, Me.cboHigherMoRun, GetWhereClause
End Sub

Private Sub cboHigherMoPart_LostFocus()
'   FillMoPartInfo cboHigherMoPart, lblHigherDescription, lblHigherType
'   FillMoRunCombo cboHigherMoPart, Me.cboHigherMoRun, GetWhereClause
End Sub

Private Sub cboHigherMoRun_Click()
   FillRunInfo cboHigherMoPart, cboHigherMoRun, lblHigherStatus, lblHigherQty
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub


Private Sub cmdPrePick_Click()
   LinkToParent
   Unload Me
End Sub

Private Sub Form_Activate()
   If bLoading Then
      bLoading = False
      FillMoPartCombo Me.cboHigherMoPart, cboHigherMoRun, GetWhereClause
      'txtPrePickQty = Me.lblLowerMoQty
   End If
End Sub

Private Sub Form_Load()
   bLoading = True
End Sub

Private Function GetWhereClause() As String
'   GetWhereClause = "where RUNSTATUS in ( 'PL', 'PP', 'PC' ) and PALEVEL <= 3" & vbCrLf _
'      & "and RUNREF <> '" & Compress(lblLowerMoPart) & "'"
   GetWhereClause = "where RUNSTATUS not like 'C%'" & vbCrLf _
      & "AND RUNREF <> '" & Compress(lblLowerMoPart) & "'"
End Function


Private Sub LinkToParent()
   On Error GoTo whoops
   Dim MoLowerRef As String, MoLowerRun As Integer
   MoLowerRef = Compress(Me.lblLowerMoPart)
   MoLowerRun = Me.lblLowerMoRun
   
   Dim MoHigherRef As String, MoHigherRun As Integer
   MoHigherRef = Compress(cboHigherMoPart)
   MoHigherRun = CLng(cboHigherMoRun)
   
   'error check
   If MoLowerRef = MoHigherRef And MoLowerRun = MoHigherRun Then
      MsgBox "You cannot link an MO to itself"
      Exit Sub
   End If
   
   Dim rs As ADODB.Recordset
   Dim bResponse As Byte

   ' confirm override
   sSql = "select 'x' from RunsTable" & vbCrLf _
      & "where RUNREF = '" & MoLowerRef & "' and RUNNO = " & MoLowerRun _
      & " and RunParentRunNo <> 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   If bSqlRows Then
      bResponse = MsgBox("MO already linked to a parent.  Override it?", ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         Exit Sub
      End If
   End If
   rs.Close
   
   ' attempt to perform the insert.  If a nonblank field is returned, it is an error message
   Dim rs2 As ADODB.Recordset
   sSql = "exec InsertMultilevelMo '" & MoHigherRef & "', " & MoHigherRun & ", '" & MoLowerRef & "', " & MoLowerRun
   bSqlRows = clsADOCon.GetDataSet(sSql, rs2)
   'If bSqlRows And rs2.Fields(0) <> "" Then
   If bSqlRows Then
      MsgBox rs2.Fields(0)
   Else
      MsgBox "MO has been linked to a parent MO."
   End If
   'rs2.Close
   Set rs2 = Nothing
   
   Exit Sub
   
whoops:
   sProcName = "LinkToParent"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
 
End Sub
