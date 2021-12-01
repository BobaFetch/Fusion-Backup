VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SysData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Databases"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "sysData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "L&og On"
      Height          =   405
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Log On To The Selected Database"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "C&ancel"
      Height          =   405
      Left            =   2880
      TabIndex        =   2
      ToolTipText     =   "Cancel The Operation And Exit"
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox cmbDbs 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Available DataBases (Caution: Includes All But System Databases (Not All Listed May Be Configured)"
      Top             =   960
      Width           =   2295
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   1800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2100
      FormDesignWidth =   4680
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
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4095
   End
   Begin VB.Label CurrDb 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   600
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
      Top             =   600
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
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "SysData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
   
End Sub

Private Sub cmbDbs_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   For iList = 0 To cmbDbs.ListCount - 1
      If cmbDbs.List(iList) = cmbDbs Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbDbs = CurrDb
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdOk_Click()
   If Trim(CurrDb) <> Trim(cmbDbs) Then LogOnToNewDb
   
End Sub

Private Sub Form_Activate()
   If bOpen = 1 Then
      FillDataBases
      bOpen = 0
   End If
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   If iBarOnTop Then
      Move MdiSect.Width / 5, MdiSect.Height / 5
   Else
      Move MdiSect.Width / 6, MdiSect.Height / 5
   End If
   FormatControls
   bOpen = 1
   Show
   
End Sub



'1/26/04

Public Sub FillDataBases()
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
      cmbDbs.enabled = False
   Else
      cmbDbs.enabled = True
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



Public Sub LogOnToNewDb()
   Dim bResponse As Byte
   Dim sNewDb As String
   Dim strConStr As String
   
   bResponse = MsgBox("You Have Selected To Change Database. Continue?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      CancelTrans
      Exit Sub
   End If
   sNewDb = cmbDbs
   On Error Resume Next
   Set clsADOCon = Nothing
'   RdoCon.Close
   Err = 0
   'RdoEnv.CursorDriver = rdUseIfNeeded
   Set clsADOCon = New ClassFusionADO
   strConStr = "Driver={SQL Server};Provider='sqloledb';UID=" & sSaAdmin & ";PWD=" & _
            sSaPassword & ";SERVER=" & sserver & ";DATABASE=" & sNewDb & ";"
   'Set RdoCon = RdoEnv.OpenConnection(dsName:="", _
   '             Prompt:=rdDriverNoPrompt, _
   '             Connect:="uid=" & sSaAdmin & ";pwd=" & sSaPassword & ";driver={SQL Server};" _
   '             & "server=" & sserver & ";database=" & sNewDb & ";")
   clsADOCon.OpenConnection (strConStr)
   
   If Err > 0 Then
      MsgBox "Couldn't Log On To The Selected Database.", _
         vbInformation, Caption
      strConStr = "Driver={SQL Server};Provider='sqloledb';UID=" & sSaAdmin & ";PWD=" & _
            sSaPassword & ";SERVER=" & sserver & ";DATABASE=" & sDataBase & ";"
      clsADOCon.OpenConnection (strConStr)
      Err = 0
      Exit Sub
   Else
      sSql = "Qry_FillCustomers"
      clsADOCon.ExecuteSQL sSql
      If Err > 0 Then
         MsgBox "The Selected Database Isn't Configured For ES/2004 ERP.", _
            vbInformation, Caption
         
      strConStr = "Driver={SQL Server};Provider='sqloledb';UID=" & sSaAdmin & ";PWD=" & _
            sSaPassword & ";SERVER=" & sserver & ";DATABASE=" & sDataBase & ";"
      clsADOCon.OpenConnection (strConStr)
         Err = 0
         
         Exit Sub
      End If
   End If
   sDataBase = sNewDb
   MdiSect.Caption = sProgName & " - " & sDataBase
   MsgBox "You Are Now Logged On To " & sDataBase & ".", _
      vbInformation, Caption
   Unload Me
   
End Sub
