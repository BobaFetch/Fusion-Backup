VERSION 5.00
Begin VB.Form ShopSHe02h 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pre-Pick to a higher level MO"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPrePickQty 
      Height          =   285
      Left            =   5340
      TabIndex        =   23
      Top             =   1620
      Width           =   795
   End
   Begin VB.ComboBox cboHigherMoPart 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   840
      Width           =   3545
   End
   Begin VB.ComboBox cboHigherMoRun 
      Height          =   315
      Left            =   5340
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select Run Number"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrePick 
      Caption         =   "Pre-Pick"
      Default         =   -1  'True
      Height          =   435
      Left            =   2160
      TabIndex        =   2
      Top             =   2580
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   435
      Left            =   3840
      TabIndex        =   3
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Label Label6 
      Caption         =   "Pre-Pick Qty"
      Height          =   255
      Left            =   4260
      TabIndex        =   22
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Type"
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblLowerType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   20
      Top             =   420
      Width           =   375
   End
   Begin VB.Label lblHigherQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5340
      TabIndex        =   19
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Qty"
      Height          =   255
      Left            =   4860
      TabIndex        =   18
      Top             =   1260
      Width           =   315
   End
   Begin VB.Label lblHigherType 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   17
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label lblLowerMoDescription 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1260
      TabIndex        =   15
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Pre-Pick to"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   900
      Width           =   975
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   1
      Left            =   4860
      TabIndex        =   13
      Top             =   900
      Width           =   435
   End
   Begin VB.Label lblHigherStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblHigherDescription 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1260
      TabIndex        =   11
      Top             =   1260
      Width           =   3495
   End
   Begin VB.Label lblLowerMoStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6660
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Qty"
      Height          =   255
      Left            =   4860
      TabIndex        =   9
      Top             =   480
      Width           =   315
   End
   Begin VB.Label lblLowerMoQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5340
      TabIndex        =   8
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Run"
      Height          =   255
      Left            =   4860
      TabIndex        =   7
      Top             =   180
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Lower MO"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   915
   End
   Begin VB.Label lblLowerMoRun 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5340
      TabIndex        =   5
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lblLowerMoPart 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1260
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "ShopSHe02h"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bLoading As Boolean
'Private whereClause As String

Private Sub cboHigherMoPart_Click()
   FillMoPartInfo cboHigherMoPart, lblHigherDescription, lblHigherType
   FillMoRunCombo cboHigherMoPart, Me.cboHigherMoRun, GetWhereClause
End Sub

Private Sub cboHigherMoRun_Click()
   FillRunInfo cboHigherMoPart, cboHigherMoRun, lblHigherStatus, lblHigherQty
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub


Private Sub cmdPrePick_Click()
   'make sure quantity is valid
   Dim moQty As Currency, PickQty As Currency
   moQty = CCur("0" & lblLowerMoQty)
   PickQty = CCur("0" & txtPrePickQty)
   If moQty <= 0 Then
      MsgBox "Pre-pick quantity required"
      Exit Sub
   End If
   
   If PickQty > moQty Then
      MsgBox "You cannot pre-pick a larger quantity than the MO quantity"
      Exit Sub
   End If
   
   PrePickMO PickQty
   MsgBox "MO has been pre-picked"
   Unload Me
End Sub

Private Sub Form_Activate()
   If bLoading Then
      bLoading = False
      FillMoPartCombo Me.cboHigherMoPart, cboHigherMoRun, GetWhereClause
      txtPrePickQty = Me.lblLowerMoQty
   End If
End Sub

Private Sub Form_Load()
   bLoading = True
End Sub

Private Function GetWhereClause() As String
   GetWhereClause = "where RUNSTATUS in ( 'PL', 'PP', 'PC' ) and PALEVEL <= 3" & vbCrLf _
      & "and RUNREF <> '" & Compress(lblLowerMoPart) & "'"
End Function


Private Sub PrePickMO(CompletionQty As Currency)
   
   'first, partially complete the lower MO
   MouseCursor ccHourglass
   On Error GoTo whoops
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   Dim bPrePick As Boolean
   Dim mo As New ClassMO
   mo.PartNumber = Me.lblLowerMoPart
   mo.RunNumber = Me.lblLowerMoRun
   bPrePick = True
   MouseCursor ccHourglass
   
   mo.CompleteMo False, CompletionQty, CDate(Format(Now, "mm/dd/yy")), bPrePick
   mo.UpdatePrePickMO cboHigherMoPart, CLng(cboHigherMoRun)
   'Set mo = Nothing
   
   'now pick the partial completion to the higher MO
   Dim pick As New ClassPick
   pick.MoPartNumber = cboHigherMoPart
   pick.MoRunNumber = CLng(cboHigherMoRun)
   
   'specify lot to use
   pick.ClearLotSelections
   pick.AddLotSelection mo.PartNumber, mo.lotNumber, CompletionQty
   
   If Not pick.PickPart(mo.PartNumber, CompletionQty, CompletionQty, _
      "?", -1, -1, "Pre", True) Then
      
      'transaction failed
      clsADOCon.RollbackTrans
      MouseCursor ccDefault
      MsgBox "Could Not Successfully Complete The Pick.  The transaction has been cancelled.", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   
   MouseCursor ccDefault
   clsADOCon.CommitTrans
'
'   If CheckLotStatus <> 0 And bPartLot And cComQty > 0 Then
'      MsgBox "The MO was completed and lot number " & vbCrLf _
'         & mo.LotNumber & " was created." & vbCrLf _
'         & "You may now edit the lot.", _
'         vbInformation, Caption
'
'      LotEdit.lblNumber = mo.LotNumber
'      LotEdit.ReadExistingMoData
'      LotEdit.Show vbModal
'   End If
'
'   On Error Resume Next
'   cmbRun.Clear
'   MouseCursor 0
'   cmbPrt.SetFocus

   Exit Sub

whoops:
   sProcName = "PrePickMO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   clsADOCon.RollbackTrans
   DoModuleErrors Me

End Sub
