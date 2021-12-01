VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CommCOe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salespersons"
   ClientHeight    =   7665
   ClientLeft      =   2370
   ClientTop       =   540
   ClientWidth     =   5640
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   315
      Left            =   3720
      TabIndex        =   84
      ToolTipText     =   "Cancel Selected Invoice"
      Top             =   5280
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   3120
      TabIndex        =   83
      Top             =   4800
      Width           =   1035
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   80
      ToolTipText     =   "Contains Customers With Invoices"
      Top             =   4800
      Width           =   1555
   End
   Begin VB.ListBox lstSelCus 
      Height          =   2010
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   79
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Frame tabFrame 
      Height          =   4092
      Index           =   1
      Left            =   5760
      TabIndex        =   63
      Top             =   1320
      Width           =   5376
      Begin VB.TextBox txtP03 
         Height          =   285
         Left            =   4200
         TabIndex        =   19
         Tag             =   "1"
         Top             =   1200
         Width           =   875
      End
      Begin VB.TextBox txtB03 
         Height          =   285
         Left            =   3240
         TabIndex        =   18
         Tag             =   "1"
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtT03 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Tag             =   "1"
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtF03 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Tag             =   "1"
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtP02 
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Tag             =   "1"
         Top             =   840
         Width           =   875
      End
      Begin VB.TextBox txtB02 
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Tag             =   "1"
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox txtT02 
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Tag             =   "1"
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox txtF02 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Tag             =   "1"
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox txtP01 
         Height          =   285
         Left            =   4200
         TabIndex        =   11
         Tag             =   "1"
         Top             =   480
         Width           =   875
      End
      Begin VB.TextBox txtB01 
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Tag             =   "1"
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox txtT01 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Tag             =   "1"
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox txtF01 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Tag             =   "1"
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox txtP10 
         Height          =   285
         Left            =   4200
         TabIndex        =   47
         Tag             =   "1"
         Top             =   3720
         Width           =   875
      End
      Begin VB.TextBox txtB10 
         Height          =   285
         Left            =   3240
         TabIndex        =   46
         Tag             =   "1"
         Top             =   3720
         Width           =   915
      End
      Begin VB.TextBox txtT10 
         Height          =   285
         Left            =   2280
         TabIndex        =   45
         Tag             =   "1"
         Top             =   3720
         Width           =   915
      End
      Begin VB.TextBox txtF10 
         Height          =   285
         Left            =   1320
         TabIndex        =   44
         Tag             =   "1"
         Top             =   3720
         Width           =   915
      End
      Begin VB.TextBox txtP09 
         Height          =   285
         Left            =   4200
         TabIndex        =   43
         Tag             =   "1"
         Top             =   3360
         Width           =   875
      End
      Begin VB.TextBox txtB09 
         Height          =   285
         Left            =   3240
         TabIndex        =   42
         Tag             =   "1"
         Top             =   3360
         Width           =   915
      End
      Begin VB.TextBox txtT09 
         Height          =   285
         Left            =   2280
         TabIndex        =   41
         Tag             =   "1"
         Top             =   3360
         Width           =   915
      End
      Begin VB.TextBox txtF09 
         Height          =   285
         Left            =   1320
         TabIndex        =   40
         Tag             =   "1"
         Top             =   3360
         Width           =   915
      End
      Begin VB.TextBox txtP08 
         Height          =   285
         Left            =   4200
         TabIndex        =   39
         Tag             =   "1"
         Top             =   3000
         Width           =   875
      End
      Begin VB.TextBox txtB08 
         Height          =   285
         Left            =   3240
         TabIndex        =   38
         Tag             =   "1"
         Top             =   3000
         Width           =   915
      End
      Begin VB.TextBox txtT08 
         Height          =   285
         Left            =   2280
         TabIndex        =   37
         Tag             =   "1"
         Top             =   3000
         Width           =   915
      End
      Begin VB.TextBox txtF08 
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Tag             =   "1"
         Top             =   3000
         Width           =   915
      End
      Begin VB.TextBox txtP07 
         Height          =   285
         Left            =   4200
         TabIndex        =   35
         Tag             =   "1"
         Top             =   2640
         Width           =   875
      End
      Begin VB.TextBox txtB07 
         Height          =   285
         Left            =   3240
         TabIndex        =   34
         Tag             =   "1"
         Top             =   2640
         Width           =   915
      End
      Begin VB.TextBox txtT07 
         Height          =   285
         Left            =   2280
         TabIndex        =   33
         Tag             =   "1"
         Top             =   2640
         Width           =   915
      End
      Begin VB.TextBox txtF07 
         Height          =   285
         Left            =   1320
         TabIndex        =   32
         Tag             =   "1"
         Top             =   2640
         Width           =   915
      End
      Begin VB.TextBox txtP06 
         Height          =   285
         Left            =   4200
         TabIndex        =   31
         Tag             =   "1"
         Top             =   2280
         Width           =   875
      End
      Begin VB.TextBox txtB06 
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Tag             =   "1"
         Top             =   2280
         Width           =   915
      End
      Begin VB.TextBox txtT06 
         Height          =   285
         Left            =   2280
         TabIndex        =   29
         Tag             =   "1"
         Top             =   2280
         Width           =   915
      End
      Begin VB.TextBox txtF06 
         Height          =   285
         Left            =   1320
         TabIndex        =   28
         Tag             =   "1"
         Top             =   2280
         Width           =   915
      End
      Begin VB.TextBox txtP05 
         Height          =   285
         Left            =   4200
         TabIndex        =   27
         Tag             =   "1"
         Top             =   1920
         Width           =   875
      End
      Begin VB.TextBox txtB05 
         Height          =   285
         Left            =   3240
         TabIndex        =   26
         Tag             =   "1"
         Top             =   1920
         Width           =   915
      End
      Begin VB.TextBox txtT05 
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Tag             =   "1"
         Top             =   1920
         Width           =   915
      End
      Begin VB.TextBox txtF05 
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         Tag             =   "1"
         Top             =   1920
         Width           =   915
      End
      Begin VB.TextBox txtP04 
         Height          =   285
         Left            =   4200
         TabIndex        =   23
         Tag             =   "1"
         Top             =   1560
         Width           =   875
      End
      Begin VB.TextBox txtB04 
         Height          =   285
         Left            =   3240
         TabIndex        =   22
         Tag             =   "1"
         Top             =   1560
         Width           =   915
      End
      Begin VB.TextBox txtT04 
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Tag             =   "1"
         Top             =   1560
         Width           =   915
      End
      Begin VB.TextBox txtF04 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Tag             =   "1"
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   17
         Left            =   4200
         TabIndex        =   77
         Top             =   240
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Base            "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   16
         Left            =   3240
         TabIndex        =   76
         Top             =   240
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Though         "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   15
         Left            =   2280
         TabIndex        =   75
         Top             =   240
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "From            "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   14
         Left            =   1320
         TabIndex        =   74
         Top             =   240
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 3"
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   73
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 2"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   72
         Top             =   840
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 1"
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 10"
         Height          =   252
         Index           =   20
         Left            =   120
         TabIndex        =   70
         Top             =   3720
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 9"
         Height          =   252
         Index           =   19
         Left            =   120
         TabIndex        =   69
         Top             =   3360
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 8"
         Height          =   252
         Index           =   18
         Left            =   120
         TabIndex        =   68
         Top             =   3000
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 7"
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   67
         Top             =   2640
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 6"
         Height          =   252
         Index           =   12
         Left            =   120
         TabIndex        =   66
         Top             =   2280
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 5"
         Height          =   252
         Index           =   11
         Left            =   120
         TabIndex        =   65
         Top             =   1920
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate 4"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   64
         Top             =   1560
         Width           =   852
      End
   End
   Begin VB.Frame tabFrame 
      Height          =   6135
      Index           =   0
      Left            =   90
      TabIndex        =   54
      Top             =   1320
      Width           =   5325
      Begin VB.ComboBox cmbAct 
         Height          =   288
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "3"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtNte 
         Height          =   855
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "9"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.ComboBox cmbReg 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   4
         Tag             =   "8"
         ToolTipText     =   "Select Region From List"
         Top             =   960
         Width           =   780
      End
      Begin VB.TextBox txtLst 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Tag             =   "2"
         Top             =   600
         Width           =   2565
      End
      Begin VB.TextBox txtMid 
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Tag             =   "3"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtFst 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Tag             =   "2"
         Top             =   240
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtPhn 
         Height          =   288
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   1452
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   285
         Index           =   24
         Left            =   120
         TabIndex        =   82
         Top             =   3480
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "More >>>"
         Height          =   252
         Left            =   4440
         TabIndex        =   78
         Top             =   240
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label lblDsc 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1320
         TabIndex        =   62
         Top             =   3000
         Width           =   2772
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         Height          =   288
         Index           =   22
         Left            =   120
         TabIndex        =   61
         Top             =   2640
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         Height          =   288
         Index           =   5
         Left            =   120
         TabIndex        =   60
         Top             =   1680
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   288
         Index           =   6
         Left            =   120
         TabIndex        =   59
         Top             =   1320
         Width           =   1152
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   58
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   57
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Init"
         Height          =   252
         Index           =   2
         Left            =   2760
         TabIndex        =   56
         Top             =   240
         Width           =   852
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1212
      End
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   6615
      Left            =   15
      TabIndex        =   53
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11668
      TabWidthStyle   =   2
      TabFixedWidth   =   1940
      TabFixedHeight  =   473
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Sales Person"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Rates"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CommCOe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optSlp 
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4680
      TabIndex        =   48
      Top             =   600
      Value           =   1  'Checked
      Width           =   252
   End
   Begin VB.ComboBox cmbSlp 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter/Revise A Salesperson (4 Char)"
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4080
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7665
      FormDesignWidth =   5640
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
      Height          =   285
      Index           =   23
      Left            =   240
      TabIndex        =   81
      Top             =   4800
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Force Use Of Salespersons"
      Enabled         =   0   'False
      Height          =   255
      Index           =   21
      Left            =   2520
      TabIndex        =   51
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salesperson"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   49
      Top             =   540
      Width           =   1215
   End
End
Attribute VB_Name = "CommCOe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of                     ***
'*** ESI Software Engineering Inc, Seattle, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

' Created: (cjs)
' Revisions:
'   08/28/03 (nth) Added vendor and accounts fields to support new commissions
'                  Group
'8/10/06 Replaced Tab with TabStrip
Option Explicit
Dim rdoSlp As ADODB.Recordset
Dim bOnLoad As Byte
Dim bGoodSprs As Byte
Dim bDataChanged As Byte

Dim sOldSprs As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbAct_Click()
   FindAccount Me
   
End Sub

Private Sub cmbAct_LostFocus()
   cmbAct = CheckLen(cmbAct, 12)
   FindAccount Me
   If bGoodSprs Then
      If Left(cmbAct, 3) <> "***" Then
         On Error Resume Next
         'rdoSlp.Edit
         rdoSlp!SPACCOUNT = Compress(cmbAct)
         rdoSlp.Update
         If Err > 0 Then ValidateEdit
      End If
   End If
End Sub

Private Sub cmbReg_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub cmbReg_LostFocus()
   cmbReg = CheckLen(cmbReg, 2)
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPREGION = "" & cmbReg
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbSlp_Click()
   bGoodSprs = GetSalesPerson(True)
End Sub

Private Sub cmbSlp_LostFocus()
   cmbSlp = CheckLen(cmbSlp, 4)
   If Len(cmbSlp) Then
      cmbSlp = Compress(cmbSlp)
      bGoodSprs = GetSalesPerson(False)
      If Not bGoodSprs Then AddSalesPerson
   Else
      bGoodSprs = False
   End If
End Sub

Private Sub cmdAdd_Click()
   Dim sItem As String
   
   Dim I As Integer
   Dim strCus As String
   Dim strSPr As String
   On Error Resume Next
   
   strSPr = Compress(cmbSlp)
   strCus = Compress(cmbCst)
   
   If (CheckIfCusExists(strCus) <> "") Then
      MsgBox "The Customer already exists in the List - " & strCus & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   
   ' Insert the part
   sSql = "INSERT INTO SprCusTable (SPCUSNUM, CUREF) VALUES('" & strSPr & "','" & strCus & "')"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   lstSelCus.AddItem strCus
   
   Exit Sub
DiaErr1:
   sProcName = "cmdAdd_Click"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdDel_Click()
   Dim sItem As String
   Dim strSPr As String
   
   Dim I As Integer
   strSPr = Compress(cmbSlp)
   With lstSelCus
      I = .ListIndex
      If I > -1 Then
         sItem = .List(I)
         On Error Resume Next
         sSql = "DELETE FROM SprCusTable WHERE SPCUSNUM = '" & strSPr & "' AND CUREF = '" & sItem & "'"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         .RemoveItem (I)
         If I = .ListCount Then
            I = I - 1
         End If
         .ListIndex = I
      End If
   End With
   
   Exit Sub
DiaErr1:
   sProcName = "cmdDel_Click"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function CheckIfCusExists(strSPr As String) As String
   Dim RdoCus As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT CUREF FROM SprCusTable WHERE SPCUSNUM = '" & strSPr & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCus, ES_FORWARD)
   If bSqlRows Then
      With RdoCus
         CheckIfCusExists = "" & Trim(!CUREF)
         ClearResultSet RdoCus
      End With
   Else
      CheckIfCusExists = ""

   End If
   Set RdoCus = Nothing
   Exit Function

modErr1:
   sProcName = "CheckIfCusExists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Function

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbSlp = ""
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2401
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillRegions
      If Trim(cUR.CurrentRegion) <> "" Then
         cmbReg = cUR.CurrentRegion
      Else
         If cmbReg.ListCount > 0 Then cmbReg = cmbReg.List(0)
      End If
      
      FillAccounts
      FillCustomers
      FillSales
      FillSelCus
      bOnLoad = 0
      bDataChanged = False
   End If
   MouseCursor 0
End Sub

Private Sub FillSelCus()
   Dim RdoSelCus As ADODB.Recordset
   Dim strSlp As String
   
   On Error GoTo modErr1
   
   strSlp = cmbSlp.Text
   
   sSql = "SELECT DISTINCT CUREF FROM SprCusTable WHERE SPCUSNUM = '" & strSlp & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSelCus, ES_FORWARD)
   If bSqlRows Then
      With RdoSelCus
         Do Until .EOF
            lstSelCus.AddItem "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet RdoSelCus
      End With
   End If
   Set RdoSelCus = Nothing
   Exit Sub

modErr1:
   sProcName = "FillSelCus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Sub


Private Sub Form_Load()
   FormLoad Me
   FormatControls
   tabFrame(0).BorderStyle = 0
   tabFrame(1).BorderStyle = 0
   tabFrame(0).Left = 40
   tabFrame(1).Left = 40
   tabFrame(1).Visible = False
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bDataChanged Then
      If Len(Trim(sOldSprs)) Then
         On Error Resume Next
         sSql = "UPDATE SprsTable SET SPREVISED='" & Format(Now, "mm/dd/yy") & "' " _
                & "WHERE SPNUMBER='" & sOldSprs & "' "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
      End If
   End If
   sSql = "UPDATE SprsTable SET SPVENDOR='NONE' WHERE SPVENDOR='' " _
          & "OR SPVENDOR IS NULL"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set rdoSlp = Nothing
   Set CommCOe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub tab1_Click()
   On Error Resume Next
   If tab1.SelectedItem.Index = 1 Then
      tabFrame(0).Visible = True
      tabFrame(1).Visible = False
      txtFst.SetFocus
   Else
      tabFrame(0).Visible = False
      tabFrame(1).Visible = True
      txtF01.SetFocus
   End If
End Sub

Private Sub TabStrip1_Click()
   
End Sub

Private Sub txtB01_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtB01_LostFocus()
   txtB01 = CheckLen(txtB01, 9)
   txtB01 = Format(Abs(Val(txtB01)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE1 = Val(txtB01)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB02_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtB02_LostFocus()
   txtB02 = CheckLen(txtB02, 9)
   txtB02 = Format(Abs(Val(txtB02)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE2 = Val(txtB02)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB03_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtB03_LostFocus()
   txtB03 = CheckLen(txtB03, 9)
   txtB03 = Format(Abs(Val(txtB03)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE3 = Val(txtB03)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB04_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtB04_LostFocus()
   txtB04 = CheckLen(txtB04, 9)
   txtB04 = Format(Abs(Val(txtB04)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE4 = Val(txtB04)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB05_LostFocus()
   txtB05 = CheckLen(txtB05, 9)
   txtB05 = Format(Abs(Val(txtB05)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE5 = Val(txtB05)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB06_LostFocus()
   txtB06 = CheckLen(txtB06, 9)
   txtB06 = Format(Abs(Val(txtB06)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE6 = Val(txtB06)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB07_LostFocus()
   txtB07 = CheckLen(txtB07, 9)
   txtB07 = Format(Abs(Val(txtB07)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE7 = Val(txtB07)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB08_LostFocus()
   txtB08 = CheckLen(txtB08, 9)
   txtB08 = Format(Abs(Val(txtB08)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE8 = Val(txtB08)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB09_LostFocus()
   txtB09 = CheckLen(txtB09, 9)
   txtB09 = Format(Abs(Val(txtB09)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE9 = Val(txtB09)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtB10_LostFocus()
   txtB10 = CheckLen(txtB10, 9)
   txtB10 = Format(Abs(Val(txtB10)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPBASE10 = Val(txtB10)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF01_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF01_LostFocus()
   txtF01 = CheckLen(txtF01, 9)
   txtF01 = Format(Abs(Val(txtF01)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM1 = Val(txtF01)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF02_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF02_LostFocus()
   txtF02 = CheckLen(txtF02, 9)
   txtF02 = Format(Abs(Val(txtF02)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM2 = Val(txtF02)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF03_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF03_LostFocus()
   txtF03 = CheckLen(txtF03, 9)
   txtF03 = Format(Abs(Val(txtF03)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM3 = Val(txtF03)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF04_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF04_LostFocus()
   txtF04 = CheckLen(txtF04, 9)
   txtF04 = Format(Abs(Val(txtF04)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM4 = Val(txtF04)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF05_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF05_LostFocus()
   txtF05 = CheckLen(txtF05, 9)
   txtF05 = Format(Abs(Val(txtF05)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM5 = Val(txtF05)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF06_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF06_LostFocus()
   txtF06 = CheckLen(txtF06, 9)
   txtF06 = Format(Abs(Val(txtF06)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM6 = Val(txtF06)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF07_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF07_LostFocus()
   txtF07 = CheckLen(txtF07, 9)
   txtF07 = Format(Abs(Val(txtF07)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM7 = Val(txtF07)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF08_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF08_LostFocus()
   txtF08 = CheckLen(txtF08, 9)
   txtF08 = Format(Abs(Val(txtF08)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM8 = Val(txtF08)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF09_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF09_LostFocus()
   txtF09 = CheckLen(txtF09, 9)
   txtF09 = Format(Abs(Val(txtF09)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM9 = Val(txtF09)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtF10_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtF10_LostFocus()
   txtF10 = CheckLen(txtF10, 9)
   txtF10 = Format(Abs(Val(txtF10)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFROM10 = Val(txtF10)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtFst_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtFst_LostFocus()
   txtFst = CheckLen(txtFst, 10)
   txtFst = StrCase(txtFst)
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPFIRST = "" & txtFst
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtLst_LostFocus()
   txtLst = CheckLen(txtLst, 20)
   txtLst = StrCase(txtLst)
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPLAST = "" & txtLst
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtMid_LostFocus()
   txtMid = CheckLen(txtMid, 1)
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPMIDD = "" & txtMid
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtNte_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtNte_LostFocus()
   txtNte = CheckLen(txtNte, 255)
   txtNte = StrCase(txtNte, ES_FIRSTWORD)
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPNOTES = "" & txtNte
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP01_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtP01_LostFocus()
   txtP01 = CheckLen(txtP01, 6)
   txtP01 = Format(Abs(Val(txtP01)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC1 = Val(txtP01)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP02_LostFocus()
   txtP02 = CheckLen(txtP02, 6)
   txtP02 = Format(Abs(Val(txtP02)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC2 = Val(txtP02)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP03_LostFocus()
   txtP03 = CheckLen(txtP03, 6)
   txtP03 = Format(Abs(Val(txtP03)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC3 = Val(txtP03)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP04_LostFocus()
   txtP04 = CheckLen(txtP04, 6)
   txtP04 = Format(Abs(Val(txtP04)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC4 = Val(txtP04)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP05_LostFocus()
   txtP05 = CheckLen(txtP05, 6)
   txtP05 = Format(Abs(Val(txtP05)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC5 = Val(txtP05)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP06_LostFocus()
   txtP06 = CheckLen(txtP06, 6)
   txtP06 = Format(Abs(Val(txtP06)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC6 = Val(txtP06)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP07_LostFocus()
   txtP07 = CheckLen(txtP07, 6)
   txtP07 = Format(Abs(Val(txtP07)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC7 = Val(txtP07)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP08_LostFocus()
   txtP08 = CheckLen(txtP08, 6)
   txtP08 = Format(Abs(Val(txtP08)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC8 = Val(txtP08)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP09_LostFocus()
   txtP09 = CheckLen(txtP09, 6)
   txtP09 = Format(Abs(Val(txtP09)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC9 = Val(txtP09)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtP10_LostFocus()
   txtP10 = CheckLen(txtP10, 6)
   txtP10 = Format(Abs(Val(txtP10)), "#0.000")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPPERC10 = Val(txtP10)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtPhn_LostFocus()
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPARCODE = Val(Left(txtPhn, 3))
      rdoSlp!SPPHONE = Right(txtPhn, 8)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub FillAccounts()
   On Error GoTo DiaErr1
   sSql = "Qry_FillLowAccounts"
   LoadComboBox cmbAct
   Exit Sub
   
DiaErr1:
   sProcName = "fillaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillSales()
   Dim RdoCmb As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_FillSalesPersons"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         cmbSlp = "" & Trim(.Fields(0))
         Do Until .EOF
            cmbSlp.AddItem "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   If cmbSlp.ListCount > 0 Then
      cmbSlp = cmbSlp.List(0)
      bGoodSprs = GetSalesPerson(True)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillsales"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetSalesPerson(bSearch As Byte) As Byte
   Dim sSalesPerson As String
   
   sSalesPerson = cmbSlp
   
   If bDataChanged Then
      If Len(Trim(sOldSprs)) Then
         On Error Resume Next
         sSql = "UPDATE SprsTable SET SPREVISED='" & Format(Now, "mm/dd/yy") & "' " _
                & "WHERE SPNUMBER='" & sOldSprs & "' "
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
      End If
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM SprsTable WHERE SPNUMBER='" & sSalesPerson & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp, ES_KEYSET)
   If bSqlRows Then
      With rdoSlp
         cmbSlp = "" & Trim(.Fields(0))
         sOldSprs = "" & Trim(.Fields(0))
         txtFst = "" & Trim(!SPFIRST)
         txtMid = "" & Trim(!SPMIDD)
         txtLst = "" & Trim(!SPLAST)
         txtFst = "" & Trim(!SPFIRST)
         cmbReg = "" & Trim(!SPREGION)
         txtPhn.Mask = ""
         txtPhn = Format(!SPARCODE, "000") & "-" & Trim(!SPPHONE)
         txtPhn.Mask = "###-###-####"
         txtF01 = Format(0 + !SPFROM1, "#####0.00")
         txtT01 = Format(0 + !SPTHRU1, "#####0.00")
         txtB01 = Format(0 + !SPBASE1, "#####0.00")
         txtP01 = Format(0 + !SPPERC1, "#0.000")
         
         txtF02 = Format(0 + !SPFROM2, "#####0.00")
         txtT02 = Format(0 + !SPTHRU2, "#####0.00")
         txtB02 = Format(0 + !SPBASE2, "#####0.00")
         txtP02 = Format(0 + !SPPERC2, "#0.000")
         
         txtF03 = Format(0 + !SPFROM3, "#####0.00")
         txtT03 = Format(0 + !SPTHRU3, "#####0.00")
         txtB03 = Format(0 + !SPBASE3, "#####0.00")
         txtP03 = Format(0 + !SPPERC3, "#0.000")
         
         txtF04 = Format(0 + !SPFROM4, "#####0.00")
         txtT04 = Format(0 + !SPTHRU4, "#####0.00")
         txtB04 = Format(0 + !SPBASE4, "#####0.00")
         txtP04 = Format(0 + !SPPERC4, "#0.000")
         
         txtF05 = Format(0 + !SPFROM5, "#####0.00")
         txtT05 = Format(0 + !SPTHRU5, "#####0.00")
         txtB05 = Format(0 + !SPBASE5, "#####0.00")
         txtP05 = Format(0 + !SPPERC5, "#0.000")
         
         txtF06 = Format(0 + !SPFROM6, "#####0.00")
         txtT06 = Format(0 + !SPTHRU6, "#####0.00")
         txtB06 = Format(0 + !SPBASE6, "#####0.00")
         txtP06 = Format(0 + !SPPERC6, "#0.000")
         
         txtF07 = Format(0 + !SPFROM7, "#####0.00")
         txtT07 = Format(0 + !SPTHRU7, "#####0.00")
         txtB07 = Format(0 + !SPBASE7, "#####0.00")
         txtP07 = Format(0 + !SPPERC7, "#0.000")
         
         txtF08 = Format(0 + !SPFROM8, "#####0.00")
         txtT08 = Format(0 + !SPTHRU8, "#####0.00")
         txtB08 = Format(0 + !SPBASE8, "#####0.00")
         txtP08 = Format(0 + !SPPERC8, "#0.000")
         
         txtF09 = Format(0 + !SPFROM9, "#####0.00")
         txtT09 = Format(0 + !SPTHRU9, "#####0.00")
         txtB09 = Format(0 + !SPBASE9, "#####0.00")
         txtP09 = Format(0 + !SPPERC9, "#0.000")
         
         txtF10 = Format(0 + !SPFROM10, "#####0.00")
         txtT10 = Format(0 + !SPTHRU10, "#####0.00")
         txtB10 = Format(0 + !SPBASE10, "#####0.00")
         txtP10 = Format(0 + !SPPERC10, "#0.000")
         
         txtNte = "" & Trim(!SPNOTES)
         
         cmbAct = "" & Trim(!SPACCOUNT)
         FindAccount Me
      End With
      bDataChanged = False
      GetSalesPerson = True
   Else
      txtFst = ""
      txtMid = ""
      txtLst = ""
      txtPhn.Mask = ""
      txtPhn = ""
      txtPhn.Mask = "###-###-####"
      txtNte = ""
      txtF01 = ""
      txtT01 = ""
      txtB01 = ""
      txtP01 = ""
      
      txtF02 = ""
      txtT02 = ""
      txtB02 = ""
      txtP02 = ""
      
      txtF03 = ""
      txtT03 = ""
      txtB03 = ""
      txtP03 = ""
      
      txtF04 = ""
      txtT04 = ""
      txtB04 = ""
      txtP04 = ""
      
      txtF05 = ""
      txtT05 = ""
      txtB05 = ""
      txtP05 = ""
      
      txtF06 = ""
      txtT06 = ""
      txtB06 = ""
      txtP06 = ""
      
      txtF07 = ""
      txtT07 = ""
      txtB07 = ""
      txtP07 = ""
      
      txtF08 = ""
      txtT08 = ""
      txtB08 = ""
      txtP08 = ""
      
      txtF09 = ""
      txtT09 = ""
      txtB09 = ""
      txtP09 = ""
      
      txtF10 = ""
      txtT10 = ""
      txtB10 = ""
      txtP10 = ""
      GetSalesPerson = False
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getsalesp"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtT01_Change()
   If Not bOnLoad Then bDataChanged = True
   
End Sub

Private Sub txtT01_LostFocus()
   txtT01 = CheckLen(txtT01, 9)
   txtT01 = Format(Abs(Val(txtT01)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU1 = Val(txtT01)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT02_LostFocus()
   txtT02 = CheckLen(txtT02, 9)
   txtT02 = Format(Abs(Val(txtT02)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU2 = Val(txtT02)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT03_LostFocus()
   txtT03 = CheckLen(txtT03, 9)
   txtT03 = Format(Abs(Val(txtT03)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU3 = Val(txtT03)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT04_LostFocus()
   txtT04 = CheckLen(txtT04, 9)
   txtT04 = Format(Abs(Val(txtT04)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU4 = Val(txtT04)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT05_LostFocus()
   txtT05 = CheckLen(txtT05, 9)
   txtT05 = Format(Abs(Val(txtT05)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU5 = Val(txtT05)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT06_LostFocus()
   txtT06 = CheckLen(txtT06, 9)
   txtT06 = Format(Abs(Val(txtT06)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU6 = Val(txtT06)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT07_LostFocus()
   txtT07 = CheckLen(txtT07, 9)
   txtT07 = Format(Abs(Val(txtT07)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU7 = Val(txtT07)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT08_LostFocus()
   txtT08 = CheckLen(txtT08, 9)
   txtT08 = Format(Abs(Val(txtT08)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU8 = Val(txtT08)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT09_LostFocus()
   txtT09 = CheckLen(txtT09, 9)
   txtT09 = Format(Abs(Val(txtT09)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU9 = Val(txtT09)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtT10_LostFocus()
   txtT10 = CheckLen(txtT10, 9)
   txtT10 = Format(Abs(Val(txtT10)), "#####0.00")
   If bGoodSprs Then
      On Error Resume Next
      'rdoSlp.Edit
      rdoSlp!SPTHRU10 = Val(txtT10)
      rdoSlp.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub AddSalesPerson()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sSalesPerson As String
   sSalesPerson = cmbSlp
   If Trim(cmbReg) = "" Then
      MsgBox "Requires At Region.", vbExclamation, Caption
      Exit Sub
   End If
   sMsg = sSalesPerson & " Wasn't Found. Add The Salesperson?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error GoTo DiaErr1
      sSql = "INSERT INTO SprsTable (SPNUMBER,SPREGION) " _
             & "VALUES('" & sSalesPerson & "','" & cmbReg & "')"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.RowsAffected Then
         SysMsg "Salesperson Added.", True
         cmbSlp = sSalesPerson
         cmbSlp.AddItem sSalesPerson
         bGoodSprs = GetSalesPerson(True)
         On Error Resume Next
         txtFst.SetFocus
      Else
         MsgBox "Couldn't Add Salesperson.", vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addsalesp"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
