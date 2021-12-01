VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form WCSchList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WorkCenter List for Schedule View"
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
   Begin VB.ListBox lstSelWC 
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
   Begin VB.ComboBox cmbWC 
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
      PictureUp       =   "WCSchList.frx":0000
      PictureDn       =   "WCSchList.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Work Centers"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1065
   End
End
Attribute VB_Name = "WCSchList"
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
' WCSchList - Assign Customer Payers
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
   Dim strWC As String
   On Error Resume Next
   
   strWC = Compress(cmbWC)
   
   If (CheckIfWCExists(strWC) <> "") Then
      MsgBox "The WC already exists in the List - " & strWC & ".", _
         vbInformation, Caption
      Exit Sub
   End If
   
   ' Insert the part
   sSql = "INSERT INTO WCSchView (WorkCenter) VALUES('" & strWC & "')"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   lstSelWC.AddItem strWC
   
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
   With lstSelWC
      I = .ListIndex
      If I > -1 Then
         sItem = .List(I)
         On Error Resume Next
         sSql = "DELETE FROM WCSchView WHERE WorkCenter = '" & sItem & "'"
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
         FillCombo
      End If
      FillSelWC
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
   Set WCSchList = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "select DISTINCT WCNREF from wcntTable"
   LoadComboBox cmbWC, -1
   Exit Sub
   
   Exit Sub
DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function CheckIfWCExists(strWC As String) As String
   Dim RdoWC As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT WorkCenter FROM WCSchView WHERE WorkCenter LIKE '" & strWC & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWC, ES_FORWARD)
   If bSqlRows Then
      With RdoWC
         CheckIfWCExists = "" & Trim(!WorkCenter)
         ClearResultSet RdoWC
      End With
   Else
      CheckIfWCExists = ""

   End If
   Set RdoWC = Nothing
   Exit Function

modErr1:
   sProcName = "CheckIfWCExists"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Function

Private Sub FillSelWC()
   Dim RdoSelWC As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT DISTINCT WorkCenter FROM WCSchView"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSelWC, ES_FORWARD)
   If bSqlRows Then
      With RdoSelWC
         Do Until .EOF
            lstSelWC.AddItem "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet RdoSelWC
      End With
   End If
   Set RdoSelWC = Nothing
   Exit Sub

modErr1:
   sProcName = "FillSelWC"
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
