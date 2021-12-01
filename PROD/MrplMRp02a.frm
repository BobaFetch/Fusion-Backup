VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form MrplMRp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP Exceptions By Part(s)"
   ClientHeight    =   4230
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkActionDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3360
      TabIndex        =   39
      Top             =   2700
      Width           =   735
   End
   Begin VB.TextBox txtPrt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   38
      Tag             =   "3"
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MrplMRp02a.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      Picture         =   "MrplMRp02a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   960
      Width           =   350
   End
   Begin VB.ComboBox cmbPart 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton ShowPrinters 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   360
      Picture         =   "MrplMRp02a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRp02a.frx":080E
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      Tag             =   "4"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame z2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   3120
      Width           =   2775
      Begin VB.OptionButton optMbe 
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   10
         Top             =   200
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   9
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   8
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   200
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbByr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "ALL"
      ToolTipText     =   "Contains Only Buyers Recorded By The MRP"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox cmbCde 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   3720
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp02a.frx":0FBC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   560
         Picture         =   "MrplMRp02a.frx":113A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4230
      FormDesignWidth =   7365
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select by action date instead of MRP date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   40
      Top             =   2700
      Width           =   3285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5640
      TabIndex        =   36
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   35
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   32
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   31
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   30
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   29
      Top             =   2220
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   28
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy, Either"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   27
      Top             =   3240
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   26
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   25
      Top             =   1860
      Width           =   1005
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   22
      Top             =   3720
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last MRP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   975
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   20
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblMrp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblUsr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   360
      Width           =   615
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   1005
      Width           =   1425
   End
End
Attribute VB_Name = "MrplMRp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/19/06 Revised report and selections. Removed extra report.
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Least to greatest dates 10/12/01

Private Sub GetMRPDates()
   Dim RdoDte As ADODB.Recordset
    sSql = "SELECT MIN(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtBeg.ToolTipText = "Earliest Date By Default"
   
    sSql = "SELECT MAX(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not IsNull(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtEnd.ToolTipText = "Latest Date By Default"
   Set RdoDte = Nothing
End Sub



Private Sub cmbByr_LostFocus()
   cmbByr = CheckLen(cmbByr, 20)
   'If Trim(cmbByr) = "" Then cmbByr = cmbByr.List(0)
   If Trim(cmbByr) = "" Then cmbByr = "ALL"
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmbCls_LostFocus()
   cmbCls = CheckLen(cmbCls, 6)
   If cmbCls = "" Then cmbCls = "ALL"
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

'Private Sub cmdFnd_Click()
'   ViewParts.lblControl = "TXTPRT"
'   ViewParts.txtPrt = txtPrt
'   optVew.Value = vbChecked
'   ViewParts.Show
'
'End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub


Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetLastMrp
      GetMRPDates
      FillBuyers
      GetOptions
      cmbCde.AddItem "ALL"
      FillProductCodes
      If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
      cmbCls.AddItem "ALL"
      FillProductClasses
      If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
      
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillCombos
      
      bOnLoad = 0
   End If
   If optVew.Value = vbChecked Then
      optVew.Value = vbUnchecked
      Unload ViewParts
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MrplMRp02a = Nothing
   
End Sub




Private Sub PrintReport()
    Dim sParts As String
    Dim sCode As String
    Dim sClass As String
    Dim sBuyer As String
    Dim sMbe As String
    Dim sBDate As String
    Dim sEDate As String
    Dim sBegDate As String
    Dim sEndDate As String
   
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    Dim strIncludes As String
    Dim strDateDev As String
   
    MouseCursor 13
    On Error GoTo DiaErr1
    GetMRPCreateDates sBegDate, sEndDate
    
    If Trim(txtBeg) = "" Then txtBeg = "ALL"
    If Trim(txtEnd) = "" Then txtEnd = "ALL"
    If Not IsDate(txtBeg) Then
       sBDate = "2000,01,01"
    Else
       sBDate = Format(txtBeg, "yyyy,mm,dd")
    End If
    If Not IsDate(txtEnd) Then
       sEDate = "2024,12,31"
    Else
       sEDate = Format(txtEnd, "yyyy,mm,dd")
    End If
    
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
    If Trim(cmbCde) = "" Then cmbCde = "ALL"
    If Trim(cmbCls) = "" Then cmbCls = "ALL"
    If Trim(cmbByr) = "" Then cmbByr = "ALL"
    If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
    If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
    If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)
    If Trim(cmbByr) = "ALL" Then sBuyer = "" Else sBuyer = Trim(cmbByr)
   
    'get custom report name if one has been defined
    sCustomReport = GetCustomReport("prdmr02")

    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
    
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "Buyer"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "DateDeveloped"
    aFormulaName.Add "Mbe"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    strIncludes = Trim(cmbPart) & ", Prod Code(s) " & cmbCde & ", Class(es) " _
                            & cmbCls
    aFormulaValue.Add CStr("'" & CStr(strIncludes) & "...'")

    aFormulaValue.Add CStr("'" & CStr(cmbByr) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")

    strDateDev = "'MRP Created  " & sBegDate & " For Requirements Through " & sEndDate & "'"
    aFormulaValue.Add CStr(strDateDev)
       
       
'    sSql = "{MrplTable.MRP_PARTREF} LIKE '" & sParts & "*' " _
'          & "AND {MrplTable.MRP_PARTPRODCODE} LIKE '" & sCode _
'          & "*' AND {MrplTable.MRP_PARTCLASS} LIKE '" & sClass & "*' " _
'          & "AND {MrplTable.MRP_POBUYER} LIKE '" & sBuyer & "*' " _
'          & "AND {MrplTable.MRP_PARTDATERQD} In Date(" & sBDate & ") " _
'          & " To Date(" & sEDate & ") AND " _
'         & " {MrplTable.MRP_TYPE} in [6, 5]"

    sSql = "{MrplTable.MRP_PARTREF} LIKE '" & sParts & "*' " _
          & "AND {MrplTable.MRP_PARTPRODCODE} LIKE '" & sCode _
          & "*' AND {MrplTable.MRP_PARTCLASS} LIKE '" & sClass & "*' " _
          & "AND {MrplTable.MRP_POBUYER} LIKE '" & sBuyer & "*' " & vbCrLf
'          & "AND {MrplTable.MRP_PARTDATERQD} In Date(" & sBDate & ") " _
'          & " To Date(" & sEDate & ") AND " _
'         & " {MrplTable.MRP_TYPE} in [6, 5]"
   
   If chkActionDate.Value = vbChecked Then
       sSql = sSql & "AND {MrplTable.MRP_ActionDate}"
   Else
       sSql = sSql & "AND {MrplTable.MRP_PARTDATERQD}"
   End If
   
    sSql = sSql & " In Date(" & sBDate & ") " _
          & " To Date(" & sEDate & ") AND " _
         & " {MrplTable.MRP_TYPE} in [6, 5]"
   
    If optMbe(0).Value = True Then
       sMbe = "Make"
       sSql = sSql & "AND {PartTable.PAMAKEBUY}='M'"
    ElseIf optMbe(1).Value = True Then
       sMbe = "Buy"
       sSql = sSql & "AND {PartTable.PAMAKEBUY}='B'"
    ElseIf optMbe(2).Value = True Then
       sMbe = "Either"
       sSql = sSql & "AND {PartTable.PAMAKEBUY}='E'"
    Else
       sMbe = "Make, Buy And Either"
    End If
   
    aFormulaValue.Add CStr("'" & sMbe & "'")
       
       ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue

    ' set the report Selection
    cCRViewer.SetReportSelectionFormula (sSql)
   
    If optDsc.Value = vbUnchecked Then
        cCRViewer.SetReportSection "GroupHeaderSection2", True
    Else
        cCRViewer.SetReportSection "GroupHeaderSection2", False
    End If
    
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    
    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
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


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
'   txtPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sCode As String * 6
   Dim sClass As String * 4
   Dim sBuyer As String * 20
   sCode = cmbCde
   sClass = cmbCls
   sBuyer = cmbByr
   sOptions = sCode & sClass & sBuyer & Trim(str(Val(optDsc.Value)))
   SaveSetting "Esi2000", "EsiProd", "Prdmr02", sOptions
   SaveSetting "Esi2000", "EsiProd", "Pmr02", lblPrinter
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr02", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      cmbCde = Mid$(sOptions, 1, 6)
      cmbCls = Mid$(sOptions, 7, 4)
      cmbByr = Trim(Mid$(sOptions, 11, 20))
      optDsc.Value = Val(Mid$(sOptions, 31, 1))
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Pmr02", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub




Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDateEx(txtBeg)
   
End Sub


Private Sub txtEnd_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDateEx(txtEnd)
   
End Sub


'Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF4 Then
'      ViewParts.lblControl = "TXTPRT"
'      ViewParts.txtPrt = txtPrt
'      optVew.Value = vbChecked
'      ViewParts.Show
'   End If
'
'End Sub

''Private Sub txtPrt_LostFocus()
 '  txtPrt = CheckLen(txtPrt, 30)
 '  If Trim(txtPrt) = "" Then txtPrt = "ALL"
 '
'End Sub



Private Sub FillBuyers()
   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT MRP_POBUYER FROM MrplTable " _
'          & "WHERE MRP_POBUYER<>'' ORDER BY MRP_POBUYER"
   
   sSql = "SELECT BYREF FROM BuyrTable ORDER BY BYREF"
   
   AddComboStr cmbByr.hwnd, "ALL"
   LoadComboBox cmbByr, -1
   'If Trim(cmbByr) = "" Then cmbByr = cmbByr.List(0)
   If Trim(cmbByr) = "" Then cmbByr = "ALL"
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPart.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPart.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      optVew.Value = vbChecked
      ViewParts.Show
   End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPart = txtPrt
End Sub

Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPart = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPart
   ViewParts.Show
End Sub

