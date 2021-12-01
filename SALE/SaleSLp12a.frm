VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLp12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Shipment Analysis"
   ClientHeight    =   3060
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLp12a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtCust 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Leading Chars Or Blank For All"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "SaleSLp12a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "SaleSLp12a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   7875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Dates Are Required)"
      Height          =   288
      Index           =   4
      Left            =   4920
      TabIndex        =   12
      Top             =   1560
      Width           =   1668
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipped From"
      Height          =   288
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1428
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Leading Characters or Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   4920
      TabIndex        =   8
      Top             =   1080
      Width           =   2988
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customers"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1425
   End
End
Attribute VB_Name = "SaleSLp12a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'9/7/06 New
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2131
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLp12a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(Now, "mm/dd/yyyy")
   txtBeg = Left(txtEnd, 3) & "01" & Right(txtEnd, 5)
   txtCust = "ALL"
   
End Sub

Private Sub optDis_Click()
   SaleSLp12b.Customers = txtCust
   SaleSLp12b.lblStart = txtBeg
   SaleSLp12b.lblThrough = txtEnd
   SaleSLp12b.Show
   
End Sub


Private Sub optPrn_Click()
   SaleSLp12b.Customers = txtCust
   SaleSLp12b.lblStart = txtBeg
   SaleSLp12b.lblThrough = txtEnd
   SaleSLp12b.Show
   
End Sub




Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtCust_LostFocus()
   If txtCust = "" Then txtCust = "ALL"
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   txtEnd = CheckDateEx(txtEnd)
   
End Sub
