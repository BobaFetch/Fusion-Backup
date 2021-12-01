VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CommCOe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Commission AP Invoice"
   ClientHeight    =   5880
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5880
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "CommCOe03a.frx":0000
      DownPicture     =   "CommCOe03a.frx":04F2
      Enabled         =   0   'False
      Height          =   372
      Left            =   6410
      MaskColor       =   &H00000000&
      Picture         =   "CommCOe03a.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   5400
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "CommCOe03a.frx":0ED6
      DownPicture     =   "CommCOe03a.frx":13C8
      Enabled         =   0   'False
      Height          =   372
      Left            =   6000
      MaskColor       =   &H00000000&
      Picture         =   "CommCOe03a.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   5400
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CommCOe03a.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   855
      Left            =   3600
      TabIndex        =   8
      Tag             =   "9"
      Top             =   2400
      Width           =   3132
   End
   Begin VB.CommandButton cmbSug 
      Caption         =   "Su&ggest"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   48
      ToolTipText     =   "Recommended Invoice Number"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox txtIdt 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1320
      TabIndex        =   5
      Tag             =   "4"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox txtPdt 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3600
      TabIndex        =   6
      Tag             =   "4"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox txtDue 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1320
      TabIndex        =   7
      Tag             =   "4"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtInv 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Tag             =   "3"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.ComboBox cmbSon 
      Height          =   288
      Left            =   1560
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Sales Orders With Commissionable Items Not Invoiced"
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox optInc 
      Caption         =   "__"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Invoice This Item"
      Top             =   5040
      Width           =   600
   End
   Begin VB.CheckBox optInc 
      Caption         =   "__"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Invoice This Item"
      Top             =   4680
      Width           =   600
   End
   Begin VB.CheckBox optInc 
      Caption         =   "__"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Invoice This Item"
      Top             =   4320
      Width           =   600
   End
   Begin VB.CheckBox optInc 
      Caption         =   "__"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Invoice This Item"
      Top             =   3960
      Width           =   600
   End
   Begin VB.CheckBox optInc 
      Caption         =   "__"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Invoice This Item"
      Top             =   3600
      Width           =   600
   End
   Begin VB.ComboBox cmbSlp 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      Top             =   340
      Width           =   855
   End
   Begin VB.CommandButton cmdCnl 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   21
      ToolTipText     =   "Cancel Operation"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdPst 
      Caption         =   "&Post"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Post Invoice"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame z3 
      Height          =   30
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   6735
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      ToolTipText     =   "Select Items"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5880
      FormDesignWidth =   6915
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4320
      TabIndex        =   61
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   4320
      TabIndex        =   60
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   59
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4320
      TabIndex        =   58
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   57
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext. Price            "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   56
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblCom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   55
      ToolTipText     =   "Quantity"
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Commission        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   54
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblCom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   53
      ToolTipText     =   "Quantity"
      Top             =   3960
      Width           =   1155
   End
   Begin VB.Label lblCom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   52
      ToolTipText     =   "Quantity"
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Label lblCom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   5640
      TabIndex        =   51
      ToolTipText     =   "Quantity"
      Top             =   4680
      Width           =   1155
   End
   Begin VB.Label lblCom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   50
      ToolTipText     =   "Quantity"
      Top             =   5040
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   288
      Index           =   11
      Left            =   2520
      TabIndex        =   49
      Top             =   2400
      Width           =   1068
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   47
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date"
      Height          =   288
      Index           =   15
      Left            =   2520
      TabIndex        =   46
      Top             =   2040
      Width           =   1068
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   45
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   44
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Total"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   43
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   41
      Top             =   732
      Width           =   1332
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1320
      TabIndex        =   40
      Top             =   720
      Width           =   252
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   39
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1320
      TabIndex        =   38
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2760
      TabIndex        =   37
      Top             =   1080
      Width           =   2772
   End
   Begin VB.Label lblSlp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2280
      TabIndex        =   36
      Top             =   340
      Width           =   2772
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   35
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   34
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   33
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   32
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   31
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   30
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   29
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   28
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   27
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1455
      TabIndex        =   26
      Top             =   3360
      Width           =   2730
   End
   Begin VB.Label lblItm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   25
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   24
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   555
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   340
      Width           =   1332
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   252
      Index           =   13
      Left            =   5160
      TabIndex        =   19
      Top             =   5520
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      Height          =   252
      Index           =   14
      Left            =   4200
      TabIndex        =   18
      Top             =   5520
      Width           =   732
   End
   Begin VB.Label lblLst 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5520
      TabIndex        =   17
      Top             =   5520
      Width           =   372
   End
   Begin VB.Label lblPge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4800
      TabIndex        =   16
      Top             =   5520
      Width           =   372
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   3960
      Picture         =   "CommCOe03a.frx":255A
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   3240
      Picture         =   "CommCOe03a.frx":2A4C
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   3480
      Picture         =   "CommCOe03a.frx":2F3E
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   3720
      Picture         =   "CommCOe03a.frx":3430
      Top             =   5520
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "CommCOe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions

' Created: 08/26/03 (nth)
' Revisions:
' 09/03/03 (nth) Made revisions per DAP
' 10/27/03 (nth) fixed commission percent
' 11/11/05 (cjs) Layout
Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodSO As Byte
Dim bPage As Byte

Dim iTotalItems As Integer
Dim iIndex As Integer
Dim iCurrPage As Integer
Dim iNetDays As Integer

'Dim rdoQry As rdoQuery
Dim cmdObj As ADODB.Command
Dim sMsg As String
Dim sSlp As String

Dim vItems(1000, 5) As Variant
' 0 = Include ?
' 1 = SO Item
' 2 = SO Item Rev
' 3 = Part
' 4 = Total Commission
' 5 = Ext Price

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetSalesOrders()
   Dim RdoFill As ADODB.Recordset
   cmbSon.Enabled = True
   sSql = "SELECT DISTINCT ITSO,SMCOSO FROM SoitTable,SpcoTable WHERE " _
          & "ITSO=SMCOSO ORDER BY ITSO"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFill, ES_FORWARD)
   If bSqlRows Then
      With RdoFill
         Do Until .EOF
            cmbSon.AddItem Format(!ITSO, SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoFill
      End With
   End If
   If cmbSon.ListCount > 0 Then
      cmbSon = cmbSon.List(0)
      bGoodSO = GetSalesOrder
   End If
   Set RdoFill = Nothing
   
End Sub


Private Sub ManageBoxs(bOn As Byte)
   Dim iList As Integer
   For iList = 1 To 5
      optInc(iList).Enabled = bOn
      lblItm(iList).Enabled = bOn
      lblPrt(iList).Enabled = bOn
      lblCom(iList).Enabled = bOn
      
      optInc(iList) = vbUnchecked
      lblItm(iList) = ""
      lblPrt(iList) = ""
      lblCom(iList) = ""
      lblExt(iList) = ""
   Next
   txtInv.Enabled = bOn
   txtDue.Enabled = bOn
   txtPdt.Enabled = bOn
   txtIdt.Enabled = bOn
   txtCmt.Enabled = bOn
   cmdCnl.Enabled = bOn
   cmbSug.Enabled = bOn
End Sub

Private Sub cmbSlp_Click()
   ESICOMM.GetThisSalesPerson Me
   GetSPSOs Me
   bGoodSO = GetSalesOrder
   
End Sub

Private Sub cmbSlp_LostFocus()
   If Not bCancel Then
      ESICOMM.GetThisSalesPerson Me
      GetSPSOs Me
      bGoodSO = GetSalesOrder
   End If
End Sub

Private Sub cmbSon_Click()
   bGoodSO = GetSalesOrder
   
End Sub

Private Sub cmbSon_LostFocus()
   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
   If Not bCancel Then bGoodSO = GetSalesOrder
   
End Sub

Private Sub cmbSug_Click()
   txtInv = Format(Now, "yymmdd") & "-" & lblPre & Trim(cmbSon)
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdCnl_Click()
   ManageBoxs False
   cmdSel.Enabled = True
   cmbSlp.Enabled = True
   cmbSon.Enabled = True
   cmdPst.Enabled = False
   cmdDn.Enabled = False
   cmdUp.Enabled = False
   cmdDn.Picture = Dsdn
   cmdUp.Picture = Dsup
   lblPge = ""
   lblLst = ""
   lblTot = ""
   txtInv = ""
   cmbSlp.SetFocus
End Sub

Private Sub cmdDn_Click()
   iCurrPage = iCurrPage + 1
   If iCurrPage > Val(lblLst) - 1 Then
      iCurrPage = Val(lblLst)
      cmdDn.Enabled = False
      cmdDn.Picture = Dsdn
   Else
      cmdDn.Enabled = True
      cmdDn.Picture = Endn
   End If
   If iCurrPage > 1 Then
      cmdUp.Enabled = True
      cmdUp.Picture = Enup
   Else
      cmdUp.Enabled = False
      cmdUp.Picture = Dsup
   End If
   lblPge = iCurrPage
   GetNextGroup
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2403
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdPst_Click()
   If Trim(txtInv) = "" Then
      sMsg = "Enter An Invoice Number."
      MsgBox sMsg, vbInformation, Caption
      txtInv.SetFocus
   Else
      PostInvoice
   End If
   
End Sub

Private Sub cmdSel_Click()
   GetItems
   
End Sub

Private Sub cmdUp_Click()
   iCurrPage = iCurrPage - 1
   If iCurrPage < 2 Then
      iCurrPage = 1
      cmdUp.Enabled = False
      cmdUp.Picture = Dsup
   Else
      cmdUp.Enabled = True
      cmdUp.Picture = Enup
   End If
   If iCurrPage < Val(lblLst) Then
      cmdDn.Enabled = True
      cmdDn.Picture = Endn
   Else
      cmdDn.Enabled = False
      cmdDn.Picture = Dsdn
   End If
   lblPge = iCurrPage
   GetNextGroup
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then FillSalesPersons Me
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls

   sCurrForm = Caption
   sSql = "SELECT DISTINCT SONUMBER,SOTYPE,SOCUST," _
          & "CUNICKNAME,CUNAME FROM SohdTable,SoitTable,CustTable " _
          & "WHERE SONUMBER=ITSO AND (SOCUST=CUREF AND SONUMBER= ?)"
   'Set rdoQry = RdoCon.CreateQuery("", sSql)
   
   Set cmdObj = New ADODB.Command
   cmdObj.CommandText = sSql
   
   Dim prmObj As ADODB.Parameter
   Set prmObj = New ADODB.Parameter
   prmObj.Type = adInteger
   
   cmdObj.Parameters.Append prmObj

   bCash = InvOrCash()
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set cmdObj = Nothing
   Set CommCOe03a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Function GetSalesOrder() As Byte
   Dim RdoSon As ADODB.Recordset
'   rdoQry(0) = Val(cmbSon)
'   bSqlRows = GetQuerySet(RdoSon, rdoQry)
   
   cmdObj.Parameters(0).Value = Val(cmbSon)
   bSqlRows = clsADOCon.GetQuerySet(RdoSon, cmdObj, ES_FORWARD, True)
   
   If bSqlRows Then
      With RdoSon
         lblPre = "" & Trim(!SOTYPE)
         lblCst = "" & Trim(!CUNICKNAME)
         lblNme = "" & Trim(!CUNAME)
         ClearResultSet RdoSon
         cmdSel.Enabled = True
         GetSalesOrder = True
      End With
   Else
      GetSalesOrder = False
      cmdSel.Enabled = False
      lblPre = ""
      lblCst = ""
      lblNme = ""
   End If
   Set RdoSon = Nothing
End Function

Private Sub GetItems()
   Dim RdoItm As ADODB.Recordset
   Dim RdoVnd As ADODB.Recordset
   Dim iList As Integer
   Dim cApplied As Currency
   
   On Error GoTo DiaErr1
   Erase vItems
   If bCash = 0 Then
      sSql = "SELECT ITNUMBER,ITREV,PARTNUM,ITQTY,ITDOLLARS,SMCOPCT,SMCOAMT,ITINVOICE,SMCOGM,PASTDCOST " _
             & "FROM SpcoTable INNER JOIN SoitTable ON SMCOSO = ITSO AND " _
             & "SMCOSOIT = ITNUMBER AND SMCOITREV = ITREV INNER JOIN PartTable PARTREF = ITPART " _
             & "WHERE ITSO = " & Val(cmbSon) & " AND ITINVOICE <> 0 AND SMCOSM = '" _
             & Trim(cmbSlp) & "'"
   Else
      sSql = "SELECT DISTINCT ITNUMBER,ITREV,PARTNUM,ITQTY,ITDOLLARS,SMCOPCT,SMCOAMT,SMCOGM,PASTDCOST " _
             & "FROM SpcoTable INNER JOIN SoitTable ON SMCOSO = ITSO AND SMCOSOIT = ITNUMBER AND " _
             & "SMCOITREV = ITREV INNER JOIN CihdTable ON ITINVOICE = INVNO INNER JOIN PartTable ON ITPART = PARTREF LEFT OUTER JOIN " _
             & "SpapTable ON SMCOSO = COSO AND SMCOSOIT = COSOIT AND SMCOITREV = COSOITREV " _
             & "AND SMCOSM = COSPNUMBER WHERE ITSO = " & Val(cmbSon) & " AND INVPIF = 1 AND " _
             & "SMCOSM = '" & Trim(cmbSlp) & "' AND COAPINV IS NULL"
   End If
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      iList = 1
      ManageBoxs True
      With RdoItm
         Do Until .EOF
            vItems(iList, 0) = 0
            vItems(iList, 1) = !ITNUMBER
            vItems(iList, 2) = !ITREV
            vItems(iList, 3) = !PartNum
            If IsNull(!SMCOGM) Or !SMCOGM = 0 Then
               vItems(iList, 4) = Format(((!ITDOLLARS * !ITQTY) * (!SMCOPCT / 100)) + !SMCOAMT, CURRENCYMASK)
            Else
               vItems(iList, 4) = Format((((!ITDOLLARS - !PASTDCOST) * !ITQTY) * (!SMCOPCT / 100)) + !SMCOAMT, CURRENCYMASK)
            End If
            vItems(iList, 5) = (!ITDOLLARS * !ITQTY)
            iList = iList + 1
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
      iTotalItems = iList - 1
      
      bPage = True
      
      iCurrPage = 1
      lblLst = CInt((((iTotalItems * 100) / 100) / 5) + 0.5)
      
      For iList = 1 To 5
         If iTotalItems >= iList Then
            optInc(iList) = vItems(iList, 0)
            lblItm(iList) = vItems(iList, 1) & vItems(iList, 2)
            lblPrt(iList) = vItems(iList, 3)
            lblCom(iList) = Format(vItems(iList, 4), CURRENCYMASK)
            lblExt(iList) = Format(vItems(iList, 5), CURRENCYMASK)
         Else
            optInc(iList).Enabled = False
            lblItm(iList).Enabled = False
            lblPrt(iList).Enabled = False
            lblCom(iList).Enabled = False
            lblExt(iList).Enabled = False
         End If
      Next
      
      If Val(lblLst) > 1 Then
         cmdDn.Enabled = True
         cmdDn.Picture = Endn
      End If
      
      bPage = False
      
      cmdPst.Enabled = True
      cmdSel.Enabled = False
      cmbSlp.Enabled = False
      cmbSon.Enabled = False
      If txtIdt = "" Then
         txtIdt = Format(ES_SYSDATE, "mm/dd/yy")
      End If
      If txtPdt = "" Then
         txtPdt = Format(ES_SYSDATE, "mm/dd/yy")
      End If
      
      ' Snag the vendor net days
      sSql = "SELECT VENETDAYS FROM VndrTable WHERE VEREF = '" _
             & GetSPVendor(cmbSlp) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
      If bSqlRows Then
         iNetDays = RdoVnd.Fields(0)
      Else
         iNetDays = 0
      End If
      Set RdoVnd = Nothing
      txtDue = Format(DateAdd("d", iNetDays, CDate(txtIdt)), "mm/dd/yy")
      txtInv.SetFocus
   Else
      sMsg = "No Items Found."
      MsgBox sMsg, vbInformation, Caption
      cmbSlp.SetFocus
   End If
   Set RdoItm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetNextGroup()
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   
   bPage = True
   
   iIndex = (Val(lblPge) - 1) * 5
   For iList = 1 To 5
      If iList + iIndex <= iTotalItems Then
         optInc(iList) = vItems(iList + iIndex, 0)
         lblItm(iList) = vItems(iList + iIndex, 1) & vItems(iList + iIndex, 2)
         lblPrt(iList) = vItems(iList + iIndex, 3)
         lblCom(iList) = Format(vItems(iList + iIndex, 4), CURRENCYMASK)
         lblExt(iList) = Format(vItems(iList + iIndex, 5), CURRENCYMASK)
         
         optInc(iList).Enabled = True
         lblItm(iList).Enabled = True
         lblPrt(iList).Enabled = True
         lblCom(iList).Enabled = True
         lblExt(iList).Enabled = True
      Else
         optInc(iList) = vbUnchecked
         lblItm(iList) = ""
         lblPrt(iList) = ""
         lblCom(iList) = ""
         lblExt(iList) = ""
         
         optInc(iList).Enabled = False
         lblItm(iList).Enabled = False
         lblPrt(iList).Enabled = False
         lblCom(iList).Enabled = False
         lblExt(iList).Enabled = False
      End If
   Next
   On Error Resume Next
   bPage = False
   optInc(1).SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "getnextgro"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub optInc_Click(Index As Integer)
   If Not bPage Then
      If optInc(Index).Value = vbChecked Then
         vItems(iIndex + Index, 0) = 1
      Else
         vItems(iIndex + Index, 0) = 0
      End If
   End If
   UpdateTotals
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 1020)
   txtCmt = CheckComments(txtCmt)
   
End Sub

Private Sub txtDue_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtDue_LostFocus()
   txtDue = CheckDate(txtDue)
End Sub

Private Sub txtIdt_Change()
   On Error Resume Next
   txtDue = Format(DateAdd("d", iNetDays, CDate(txtIdt)), "mm/dd/yy")
End Sub

Private Sub txtIdt_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtIdt_LostFocus()
   txtIdt = CheckDate(txtIdt)
End Sub

Private Sub PostInvoice()
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim iResponse As Integer
   Dim iList As Integer
   Dim b As Byte
   Dim cTotal As Currency
   Dim sJournalID As String
   Dim sVendor As String
   Dim sComAcct As String
   Dim sApAcct As String
   Dim sInv As String
   Dim sNote As String
   
   On Error GoTo DiaErr1
   
   For iList = 1 To iTotalItems
      If vItems(iList, 0) = 1 Then
         b = True
         Exit For
      End If
   Next
   If Not b Then
      sMsg = "There Are No Items Selected To Be Invoiced."
      MsgBox sMsg, vbInformation, Caption
      optInc(1).SetFocus
      Exit Sub
   End If
   
   sJournalID = GetOpenJournal("PJ", Format(txtPdt, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   
   If b = 0 Then
      MsgBox "There Is No Open Purchases Journal For The Period.", _
         vbExclamation, Caption
      txtPdt.SetFocus
      Exit Sub
   End If
   
   sInv = Trim(txtInv)
   
   sMsg = "Post Invoice " & sInv
   iResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If iResponse = vbNo Then
      Exit Sub
   End If
   
   sVendor = GetSPVendor(cmbSlp)
   sComAcct = GetSPAccount(cmbSlp)
   sApAcct = GetAPAccount()
   cTotal = CCur(lblTot)
   
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "INSERT INTO VihdTable (VINO,VIVENDOR," _
          & "VIDATE,VIDUE,VIDTRECD,VIDUEDATE,VICOMT) " _
          & "VALUES('" & txtInv & "','" & sVendor & "','" _
          & txtIdt & "'," & cTotal & ",'" & txtPdt & "','" _
          & txtDue & "','" _
          & Trim(txtCmt) & "')"
   clsADOCon.ExecuteSQL sSql
   For iList = 1 To iTotalItems
      If vItems(iList, 0) = 1 Then
         
         sNote = lblPre & cmbSon & " " & vItems(iList, 1) _
                 & vItems(iList, 2) & " " & vItems(iList, 3)
         
         sSql = "INSERT INTO ViitTable (VITNO,VITVENDOR,VITITEM," _
                & "VITQTY,VITCOST,VITACCOUNT,VITNOTE) " _
                & "VALUES('" _
                & sInv & "','" _
                & sVendor & "'," _
                & iList & "," _
                & "1," _
                & CCur(vItems(iList, 4)) & ",'" _
                & sComAcct & "','" _
                & sNote & "')"
         clsADOCon.ExecuteSQL sSql
         sSql = "INSERT INTO SpapTable (COSO,COSOIT,COSOITREV,COAPINV," _
                & "COAPVENDOR,COSPNUMBER) VALUES( " _
                & Val(cmbSon) & "," _
                & vItems(iList, 1) & ",'" _
                & vItems(iList, 2) & "','" _
                & txtInv & "','" _
                & sVendor & "','" _
                & cmbSlp & "')"
         clsADOCon.ExecuteSQL sSql
         
         iTrans = GetNextTransaction(sJournalID)
         If iTrans > 0 Then
            
            ' Credit
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCCREDIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & CCur(vItems(iList, 4)) & ",'" _
                   & sApAcct & "','" _
                   & Format(Now, "mm/dd/yyyy") & "','" _
                   & sVendor & "','" _
                   & Trim(txtInv) & "')"
            clsADOCon.ExecuteSQL sSql
            
            ' Debit
            iRef = iRef + 1
            sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF," _
                   & "DCDEBIT,DCACCTNO,DCDATE,DCVENDOR,DCVENDORINV) " _
                   & "VALUES('" _
                   & Trim(sJournalID) & "'," _
                   & iTrans & "," _
                   & iRef & "," _
                   & CCur(vItems(iList, 4)) & ",'" _
                   & sComAcct & "','" _
                   & Format(Now, "mm/dd/yyyy") & "','" _
                   & sVendor & "','" _
                   & Trim(txtInv) & "')"
            clsADOCon.ExecuteSQL sSql
         End If
      End If
   Next
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      sMsg = "Successfully Posted."
      SysMsg sMsg, True
      ManageBoxs False
      cmdCnl_Click
   Else
      clsADOCon.RollbackTrans
      sMsg = "Could Not Post " & sInv & vbCrLf _
             & "Transaction Canceled."
      MsgBox sMsg, vbExclamation, Caption
   End If
   
   Exit Sub
DiaErr1:
   sProcName = "postinvoi"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtInv_LostFocus()
   txtInv = CheckLen(txtInv, 20)
   txtInv = CheckComments(txtInv)
End Sub

Private Sub txtPdt_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtPdt_LostFocus()
   txtPdt = CheckDate(txtPdt)
End Sub

Private Sub UpdateTotals()
   Dim iList As Integer
   Dim cTotal As Currency
   For iList = 1 To iTotalItems
      If vItems(iList, 0) = 1 Then
         cTotal = cTotal + CCur(vItems(iList, 4))
      End If
   Next
   lblTot = Format(cTotal, CURRENCYMASK)
   
End Sub

Private Function GetAPAccount() As String
   Dim RdoAct As ADODB.Recordset
   sSql = "SELECT COAPACCT FROM ComnTable WHERE COREF = 1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct)
   If bSqlRows Then
      With RdoAct
         GetAPAccount = "" & .Fields(0)
      End With
      ClearResultSet RdoAct
   End If
   Set RdoAct = Nothing
   
End Function
