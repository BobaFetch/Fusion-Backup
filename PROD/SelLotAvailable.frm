VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form SelLotAvailable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Lot Available"
   ClientHeight    =   5175
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5175
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Tag             =   "3"
      ToolTipText     =   "Select Part Number "
      Top             =   480
      Width           =   3300
   End
   Begin VB.ListBox lstSelLot 
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Top             =   1080
      Width           =   915
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   2160
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   6495
   End
   Begin VB.ComboBox cmbLot 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Contains selected Lots"
      Top             =   1080
      Width           =   3330
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
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
      PictureUp       =   "SelLotAvailable.frx":0000
      PictureDn       =   "SelLotAvailable.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Lot Numbers"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Available Lotnumbers"
      Height          =   405
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1185
   End
End
Attribute VB_Name = "SelLotAvailable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*********************************************************************************
' SelLotAvailable - Assign Customer Payers
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

Private Sub cmbPrt_Click()
   GetLotNumbers
End Sub

Private Sub cmbPrt_LostFocus()
    GetLotNumbers
End Sub

Private Sub cmdAdd_Click()
   Dim sItem As String
   
   Dim I As Integer
   Dim strLot As String
   On Error Resume Next
   
   strLot = cmbLot
   
   If (CheckIfLotExists(strLot) <> "") Then
      MsgBox "The Lot Number already exists in the List - " & strLot & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   
   ' Insert the part
   sSql = "INSERT INTO LtTrkTable (LOTUSERLOTID) VALUES('" & strLot & "')"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   sSql = "UPDATE a SET a.LOTNUMBER = b.LOTNUMBER FROM " _
            & "LtTrkTable a, lohdTable b WHERE a.LOTUSERLOTID  = '" & strLot & "'" _
            & " AND a.LOTUSERLOTID = b.LOTUSERLOTID"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   lstSelLot.AddItem strLot
   
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
   With lstSelLot
      I = .ListIndex
      If I > -1 Then
         sItem = .List(I)
         On Error Resume Next
         sSql = "DELETE FROM LtTrkTable WHERE LOTUSERLOTID = '" & sItem & "'"
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
         FillLotParts
         GetLotNumbers
      End If
      FillSelLot
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
   Set SelLotAvailable = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillLotParts()
   On Error GoTo modErr1
   sSql = "SELECT DISTINCT PARTNUM FROM " _
          & "PartTable,LohdTable WHERE LOTREMAININGQTY > 0 " _
          & " AND PARTREF=LOTPARTREF AND palevel = 5 ORDER BY PARTNUM"
          
   LoadComboBox cmbPrt, -1
   Exit Sub
   
modErr1:
   sProcName = "FillLotParts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Private Function CheckIfLotExists(strLot As String) As String
   Dim RdoWC As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT LOTUSERLOTID FROM LtTrkTable WHERE LOTUSERLOTID = '" & strLot & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWC, ES_FORWARD)
   If bSqlRows Then
      With RdoWC
         CheckIfLotExists = "" & Trim(!LOTUSERLOTID)
         ClearResultSet RdoWC
      End With
   Else
      CheckIfLotExists = ""

   End If
   Set RdoWC = Nothing
   Exit Function

modErr1:
   sProcName = "CheckIfLotExists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Function

Private Sub GetLotNumbers()

   On Error GoTo modErr1
   Dim strPartNum As String
   
   cmbLot.Clear
   strPartNum = Compress(cmbPrt)
   sSql = "SELECT DISTINCT lotuserlotid FROM lohdTable WHERE LOTPARTREF= '" & strPartNum & "'"
   
   LoadComboBox cmbLot, -1
   Exit Sub

modErr1:
   sProcName = "GetLotNumbers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Sub

Private Sub FillSelLot()
   Dim RdoSelWC As ADODB.Recordset
   On Error GoTo modErr1
   Dim strPartNum As String
   
   strPartNum = Compress(cmbPrt)
   sSql = "SELECT DISTINCT lotuserlotid FROM LtTrkTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSelWC, ES_FORWARD)
   If bSqlRows Then
      With RdoSelWC
         Do Until .EOF
            lstSelLot.AddItem "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet RdoSelWC
      End With
   End If
   Set RdoSelWC = Nothing
   Exit Sub

modErr1:
   sProcName = "FillSelLot"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Sub

Private Sub lstSelWC_Click()

'   Dim i As Integer
'   Dim sItem As String
'
'   i = lstAva.ListIndex
'   If i > -1 Then
'      sItem = .List(i)
'      txtPrt.Text = sItem
'   End If
   
End Sub
