VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe12a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Customer Payers"
   ClientHeight    =   3720
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3720
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSel 
      Height          =   2010
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ListBox lstAva 
      Height          =   2010
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
      Height          =   315
      Left            =   2430
      TabIndex        =   2
      Top             =   1680
      Width           =   915
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "<<"
      Height          =   315
      Left            =   2430
      TabIndex        =   3
      ToolTipText     =   "Cancel Selected Invoice"
      Top             =   2040
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   5535
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   960
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   5
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
      PictureUp       =   "diaARe12a.frx":0000
      PictureDn       =   "diaARe12a.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   11
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Availiable"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer "
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   825
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   720
      Width           =   3000
   End
End
Attribute VB_Name = "diaARe12a"
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
' diaARe12a - Assign Customer Payers
'
' Notes:
'
' Created: (nth) 07/12/04
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Public bRemote As Byte
Dim sCust As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub ManageBoxes()
   If lstAva.ListCount > 0 Then
      cmdAdd.enabled = True
   Else
      cmdAdd.enabled = False
   End If
   If lstSel.ListCount > 0 Then
      cmdDel.enabled = True
   Else
      cmdDel.enabled = False
   End If
End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
End Sub

Private Sub cmbCst_LostFocus()
   If Not bCancel Then
      FindCustomer Me, cmbCst
      sCust = Compress(cmbCst)
      GetPayers
   End If
End Sub

Private Sub cmdAdd_Click()
   Dim sItem As String
   Dim i As Integer
   With lstAva
      i = .ListIndex
      If i > -1 Then
         sItem = .List(i)
         On Error Resume Next
         sSql = "INSERT INTO CpayTable (CPCUST,CPPAYER) VALUES('" & sCust _
                & "','" & Compress(sItem) & "')"
         clsADOCon.ExecuteSQL sSql
         If Err > 0 Then
            ValidateEdit Me
         Else
            On Error GoTo DiaErr1
            lstSel.AddItem sItem
            .RemoveItem (i)
            If i = .ListCount Then
               i = i - 1
            End If
            .ListIndex = i
         End If
      End If
   End With
   ManageBoxes
   Exit Sub
DiaErr1:
   sProcName = "getpayers"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdDel_Click()
   Dim sItem As String
   Dim i As Integer
   With lstSel
      i = .ListIndex
      If i > -1 Then
         sItem = .List(i)
         On Error Resume Next
         sSql = "DELETE FROM CpayTable WHERE CPCUST = '" & sCust _
                & "' AND CPPAYER = '" & Compress(sItem) & "'"
         clsADOCon.ExecuteSQL sSql
         If Err > 0 Then
            ValidateEdit Me
         Else
            On Error GoTo DiaErr1
            lstAva.AddItem sItem
            .RemoveItem (i)
            If i = .ListCount Then
               i = i - 1
            End If
            .ListIndex = i
         End If
      End If
   End With
   ManageBoxes
   Exit Sub
DiaErr1:
   sProcName = "getpayers"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      If Not bRemote Then
         FillCombo
         cmbCst = cUR.CurrentCustomer
      Else
         cmbCst.enabled = False
      End If
      FindCustomer Me, cmbCst
      sCust = Compress(cmbCst)
      GetPayers
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
      cUR.CurrentCustomer = cmbCst
      SaveCurrentSelections
   Else
      diaARe04a.sGetInvoices
   End If
   Set diaARe12a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   FillCustomers Me
   Exit Sub
DiaErr1:
   sProcName = "fillcomb"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetPayers()
   Dim rdoCst As ADODB.Recordset
   Dim sPayers() As String
   Dim i As Integer
   Dim sIn As String
   
   lstAva.Clear
   lstSel.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT CUNICKNAME,CPPAYER FROM CpayTable INNER JOIN CustTable ON " _
          & "CPPAYER = CUREF WHERE CPCUST = '" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         While Not .EOF
            lstSel.AddItem "" & Trim(.Fields(0))
            sIn = sIn & "'" & "" & Trim(.Fields(1)) & "',"
            .MoveNext
         Wend
         .Cancel
      End With
      Set rdoCst = Nothing
      sIn = "WHERE CUREF NOT IN (" & sIn & "'" & sCust & "')"
   End If
   sSql = "SELECT CUNICKNAME FROM CustTable " & sIn
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst)
   If bSqlRows Then
      With rdoCst
         While Not .EOF
            lstAva.AddItem "" & Trim(.Fields(0))
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set rdoCst = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getpayers"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
