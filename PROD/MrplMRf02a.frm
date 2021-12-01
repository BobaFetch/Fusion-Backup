VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form MrplMRf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Lot Location to Exclude"
   ClientHeight    =   3990
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3990
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSelLoc 
      Height          =   2400
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   915
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Cancel Selected Invoice"
      Top             =   1440
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5535
   End
   Begin VB.ComboBox cmbLoc 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Contains Customers With Invoices"
      Top             =   360
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
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
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "MrplMRf02a.frx":0000
      PictureDn       =   "MrplMRf02a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Lot Location"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Location"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1065
   End
End
Attribute VB_Name = "MrplMRf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***


Option Explicit

'*********************************************************************************
' MrplMRf02a - Assign Customer Payers
'
' Notes:
'
' Created:
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Public bRemote As Byte
Dim sRpt As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmdAdd_Click()
   Dim sItem As String
   
   Dim I As Integer
   Dim strLoc As String
   On Error Resume Next
   
   strLoc = Compress(cmbLoc)
   
   If strLoc = "" Then
      MsgBox "You cannot add a blank location", _
         vbInformation, Caption
      Exit Sub
   End If
   
   If (CheckIfLocExists(strLoc) <> "") Then
      MsgBox "The Loc already exists in the List - " & strLoc & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   
   ' Insert the part
   sSql = "INSERT INTO LoLcTable (LOTEXLOCATION) VALUES('" & strLoc & "')"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   lstSelLoc.AddItem strLoc
   
   Exit Sub
DiaErr1:
   sProcName = "cmdAdd_Click"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdDel_Click()
   Dim sItem As String
   Dim I As Integer
   With lstSelLoc
      I = .ListIndex
      If I > -1 Then
         sItem = .List(I)
         On Error Resume Next

         sSql = "DELETE FROM LoLcTable WHERE LOTEXLOCATION = '" & sItem & "'"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
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

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      If Not bRemote Then
         FillCombo
      End If
      FillSelLoc
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not bRemote Then
      FormUnload
      SaveCurrentSelections
   End If
   Set MrplMRf02a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "select DISTINCT LOTLOCATION from LohdTable WHERE (LOTLOCATION IS NOT NULL OR LOTLOCATION <> '')"
   LoadComboBox cmbLoc, -1
   Exit Sub
   
   Exit Sub
DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function CheckIfLocExists(strLoc As String) As String
   Dim RdoLoc As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT LOTEXLOCATION FROM LoLcTable WHERE LOTEXLOCATION LIKE '" & strLoc & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLoc, ES_FORWARD)
   If bSqlRows Then
      With RdoLoc
         CheckIfLocExists = "" & Trim(!LOTEXLOCATION)
         ClearResultSet RdoLoc
      End With
   Else
      CheckIfLocExists = ""

   End If
   Set RdoLoc = Nothing
   Exit Function

modErr1:
   sProcName = "CheckIfLocExists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Function

Private Sub FillSelLoc()
   Dim RdoSelLoc As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT DISTINCT LOTEXLOCATION FROM LoLcTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSelLoc, ES_FORWARD)
   If bSqlRows Then
      With RdoSelLoc
         Do Until .EOF
            lstSelLoc.AddItem "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet RdoSelLoc
      End With
   End If
   Set RdoSelLoc = Nothing
   Exit Sub

modErr1:
   sProcName = "FillSelLoc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Sub


