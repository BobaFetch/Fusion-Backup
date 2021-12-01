VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SysData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Databases"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   ClipControls    =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   931
   Icon            =   "sysData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "sysData.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      ToolTipText     =   "Cancel The Operation And Exit"
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbKey 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "No User Functions"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "L&og On"
      Height          =   315
      Index           =   0
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Log On To The Selected Database"
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      ToolTipText     =   "Cancel The Operation And Exit"
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cmbDbs 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Available DataBases (Caution: Includes All But System Databases (Not All Listed May Be Configured)"
      Top             =   1200
      Width           =   2295
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3060
      FormDesignWidth =   4785
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emulation Mode"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "No User Functions"
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Change The Database In Use To Another"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   4095
   End
   Begin VB.Label CurrDb 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1995
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Database"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Databases"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "SysData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'7/13/05 Added ES_CUSTOM/cmbKey for emulating customers
'11/16/05 Changed the connection object
'10/16/06 Changed the DSN Registration
Option Explicit
Dim bOpen As Byte
Dim sOldDb As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   CurrDb.ForeColor = ES_BLUE
   CurrDb.Caption = sDataBase
   sOldDb = sDataBase
   cmbDbs.ToolTipText = "Select The Database From The List - Caution " _
                        & "A Listed Database May Not Be Configured For " & sSysCaption
   cmbKey.BackColor = Me.BackColor
   cmbKey.ToolTipText = "No User Functions"
   
End Sub

Private Sub cmbDbs_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   For iList = 0 To cmbDbs.ListCount - 1
      If cmbDbs.List(iList) = cmbDbs Then b = 1
   Next
   If b = 0 Then
      cmbDbs = CurrDb
   End If
   
End Sub


Private Sub cmbKey_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If cmbKey = "KIHEI" Then
      FillMe
   Else
      If cmbKey.ListCount > 0 Then
         For iList = 0 To cmbKey.ListCount - 1
            If cmbKey.List(iList) = cmbKey Then bByte = 1
         Next
      End If
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 931
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdOk_Click(Index As Integer)
   If Trim(CurrDb) <> Trim(cmbDbs) Then LogOnToNewDb
   
End Sub

Private Sub cmdUpd_Click()
   ES_CUSTOM = Trim(cmbKey)
   MsgBox "You Are Now Emulating " & ES_CUSTOM & ".", _
      vbInformation, Caption
   Sleep 500
   Unload Me
   
End Sub

Private Sub Form_Activate()
   If bOpen = 1 Then FillDataBases
   
End Sub

Private Sub Form_Initialize()
   CloseForms
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   If iBarOnTop Then
      Move MDISect.Width / 5, MDISect.Height / 5
   Else
      Move MDISect.Width / 6, MDISect.Height / 5
   End If
   FormatControls
   bOpen = 1
   
End Sub



'1/26/04

Private Sub FillDataBases()
   Dim b As Byte
   Dim aRow As Byte
   Dim sData(12) As String
   
   On Error Resume Next
   For b = 0 To 11
      sData(b) = Trim(GetSetting("Esi2000", "System", "UserDatabase" & Trim(str(b)), sData(b)))
      If Trim(sData(b)) <> "" Then cmbDbs.AddItem sData(b)
   Next
   If cmbDbs.ListCount = 0 Then
      cmbDbs.AddItem "Esi2000Db"
      cmbDbs.Enabled = False
   Else
      cmbDbs.Enabled = True
   End If
   
   For aRow = 0 To cmbDbs.ListCount - 1
      If Trim(cmbDbs.List(aRow)) = "Esi2000Db" Then
         b = 1
         Exit For
      Else
         b = 0
      End If
   Next
   If b = 0 Then cmbDbs.AddItem "Esi2000Db"
   cmbDbs = sDataBase
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set SysData = Nothing
   
End Sub

Private Sub ReconnectToExistingDB()
    OpenDBServer True
    MouseCursor ccArrow
    
End Sub



Private Sub LogOnToNewDb()
   Dim bResponse As Byte
   Dim sNewDb As String
   Dim sConnStr As String
     Dim ErrNum    As Long
   Dim ErrDesc   As String
   Dim bQueryOk As Boolean
 
   bResponse = MsgBox("You Have Selected To Change Database. Continue?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If
   sNewDb = cmbDbs
   On Error Resume Next
   Set clsADOCon = Nothing
   Err.Clear
   
      
    Set clsADOCon = New ClassFusionADO
   
    sConnStr = "Driver={SQL Server};Provider='sqloledb';UID=" & sSaAdmin & ";PWD=" & _
            sSaPassword & ";SERVER=" & sserver & ";DATABASE=" & sNewDb & ";"
 
   If clsADOCon.OpenConnection(sConnStr, ErrNum, ErrDesc) = False Then
      MsgBox "Couldn't Log On To The Selected Database.", _
         vbInformation, Caption
      Err.Clear
      Set clsADOCon = Nothing
      ReconnectToExistingDB
      Exit Sub
   Else
      sSql = "Qry_FillSortedCustomers"
      bQueryOk = clsADOCon.ExecuteSQL(sSql)
    
      If Not bQueryOk Then
         MsgBox "The Selected Database Isn't Configured For ES/2004 ERP.", _
            vbInformation, Caption
         Err.Clear
         Set clsADOCon = Nothing
         ReconnectToExistingDB
         Exit Sub
      End If

 
   
   End If

   sDataBase = sNewDb
 '  sDsn = RegisterSqlDsn("ESI2000")
   cUR.CurrentPart = ""
   cUR.CurrentVendor = ""
   cUR.CurrentCustomer = ""
   'MdiSect.Caption = sProgName & " - " & sDataBase
   MDISect.Caption = GetSystemCaption
   MsgBox "You Are Now Logged On To " & sDataBase & ".", _
      vbInformation, Caption
   SaveUserSetting USERSETTING_DatabaseName, sNewDb

   Unload Me
   
End Sub


Private Sub FillMe()
   cmbKey.BackColor = Es_TextBackColor
   cmbKey.AddItem "<NONE>"
   cmbKey.AddItem "WATERJET"
   cmbKey.AddItem "INTCOA"
   cmbKey.AddItem "PROPLA"
   cmbKey.AddItem "JEVCO"
   cmbKey.AddItem "PF-SBS"
   If ES_CUSTOM = "" Then
      cmbKey = cmbKey.List(0)
   Else
      cmbKey = ES_CUSTOM
   End If
   cmdUpd.Visible = True
   
End Sub



