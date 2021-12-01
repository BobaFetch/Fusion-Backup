VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHe07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Internal Comment"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtCmt 
      Height          =   2535
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   4695
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5520
      Top             =   960
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4260
      FormDesignWidth =   6090
   End
   Begin VB.CommandButton cmdFnd 
      Height          =   375
      Left            =   4680
      Picture         =   "ShopSHe07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Find Part"
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Part Number "
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
End
Attribute VB_Name = "ShopSHe07a"
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

Dim bGoodPart As Byte
Dim RdoPart As ADODB.Recordset
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Change()
     bGoodPart = GetPart
End Sub

Private Sub cmbPrt_Click()
     bGoodPart = GetPart
End Sub

Private Sub cmdCan_Click()
    Unload Me
End Sub

Private Sub cmdFnd_Click()
    ViewParts.lblControl = "txtPrt"
    ViewParts.txtPrt = txtPrt
    ViewParts.Show
    bGoodPart = GetPart
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   MouseCursor 0
   FillPartCombo cmbPrt
   cmbPrt = ""
End Sub

Private Sub Form_Load()
   sSql = "SELECT PARTREF, PADESC, PAINTERNALCMT FROM PartTable WHERE PARTREF = ? "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   
   FormLoad Me
   FormatControls
   
   txtCmt.Enabled = False
   bGoodPart = 0
End Sub

Private Sub Form_Resize()
   Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   FormUnload
   Set ShopSHe07a = Nothing
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Function GetPart() As Byte
    GetPart = 0
    
    AdoQry.Parameters(0).Value = Compress(cmbPrt)
    bSqlRows = clsADOCon.GetQuerySet(RdoPart, AdoQry, ES_KEYSET, True)
    If bSqlRows Then
        lblDsc = "" & RdoPart!PADESC
        txtCmt = "" & RdoPart!PAINTERNALCMT
        txtCmt.Enabled = True
        GetPart = 1
    Else
        lblDsc = "*** Part Not Found ***"
        txtCmt = ""
        txtCmt.Enabled = False
        GetPart = 0
    End If
End Function


Private Sub txtCmt_LostFocus()
   If bGoodPart Then
      On Error Resume Next
      RdoPart!PAINTERNALCMT = "" & txtCmt
      RdoPart.Update
   End If
End Sub


Private Sub txtPrt_Change()
    bGoodPart = GetPart
End Sub


Private Sub txtPrt_LostFocus()
    bGoodPart = GetPart
End Sub


Private Sub cmbPrt_LostFocus()
    bGoodPart = GetPart
End Sub

