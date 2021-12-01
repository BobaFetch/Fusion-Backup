VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form ShopSHe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Manufacturing Parameters"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<<&Prev"
      Height          =   255
      Left            =   5640
      TabIndex        =   53
      Top             =   4560
      Width           =   1000
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ShopSHe05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   "&Next >>>"
      Height          =   255
      Left            =   6840
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Get Part Numbers"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtLed 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   6720
      TabIndex        =   16
      ToolTipText     =   "Purchasing Lead In Days"
      Top             =   4200
      Width           =   925
   End
   Begin VB.TextBox txtFlw 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   5760
      TabIndex        =   15
      ToolTipText     =   "Manufacturing Flow In Days"
      Top             =   4200
      Width           =   925
   End
   Begin VB.TextBox txtLed 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6720
      TabIndex        =   14
      ToolTipText     =   "Purchasing Lead In Days"
      Top             =   3840
      Width           =   925
   End
   Begin VB.TextBox txtFlw 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   5760
      TabIndex        =   13
      ToolTipText     =   "Manufacturing Flow In Days"
      Top             =   3840
      Width           =   925
   End
   Begin VB.TextBox txtLed 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6720
      TabIndex        =   12
      ToolTipText     =   "Purchasing Lead In Days"
      Top             =   3480
      Width           =   925
   End
   Begin VB.TextBox txtFlw 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5760
      TabIndex        =   11
      ToolTipText     =   "Manufacturing Flow In Days"
      Top             =   3480
      Width           =   925
   End
   Begin VB.TextBox txtLed 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   6720
      TabIndex        =   10
      ToolTipText     =   "Purchasing Lead In Days"
      Top             =   3120
      Width           =   925
   End
   Begin VB.TextBox txtFlw 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   9
      ToolTipText     =   "Manufacturing Flow In Days"
      Top             =   3120
      Width           =   925
   End
   Begin VB.TextBox txtLed 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6720
      TabIndex        =   8
      ToolTipText     =   "Purchasing Lead In Days"
      Top             =   2760
      Width           =   925
   End
   Begin VB.TextBox txtFlw 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Manufacturing Flow In Days"
      Top             =   2760
      Width           =   925
   End
   Begin VB.TextBox txtLed 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   6
      ToolTipText     =   "Purchasing Lead In Days"
      Top             =   2400
      Width           =   925
   End
   Begin VB.TextBox txtFlw 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5760
      TabIndex        =   5
      ToolTipText     =   "Manufacturing Flow In Days"
      Top             =   2400
      Width           =   925
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Select"
      Height          =   315
      Left            =   6840
      TabIndex        =   4
      ToolTipText     =   "Get Part Numbers"
      Top             =   1440
      Width           =   875
   End
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   5280
      TabIndex        =   3
      Tag             =   "8"
      ToolTipText     =   "Select Level Or Type"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Enter Leading Char(s) Or Blank (200 Max Selected)"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Frame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   7572
   End
   Begin VB.TextBox txtDefFlow 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtDefLead 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6840
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4875
      FormDesignWidth =   7785
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasing"
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   8
      Left            =   6720
      TabIndex        =   52
      Top             =   1960
      Width           =   1092
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing"
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   7
      Left            =   5760
      TabIndex        =   51
      Top             =   1960
      Width           =   1092
   End
   Begin VB.Label lblLvl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   48
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   47
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   46
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblLvl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   45
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   2760
      TabIndex        =   44
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   43
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblLvl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   42
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   41
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   40
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblLvl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   39
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   38
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   37
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblLvl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   36
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   35
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   34
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   7680
      X2              =   5280
      Y1              =   2376
      Y2              =   2376
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   5160
      X2              =   2760
      Y1              =   2380
      Y2              =   2380
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2640
      X2              =   240
      Y1              =   2380
      Y2              =   2380
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lead"
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   6
      Left            =   6720
      TabIndex        =   33
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Flow"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   5760
      TabIndex        =   32
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   4
      Left            =   5280
      TabIndex        =   31
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   2760
      TabIndex        =   30
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   240
      TabIndex        =   29
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblLvl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   5280
      TabIndex        =   28
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   27
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   1
      Left            =   4560
      TabIndex        =   25
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Parts"
      ForeColor       =   &H00000000&
      Height          =   228
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Shop Floor Settings:"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   42
      Left            =   240
      TabIndex        =   22
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Purchasing Lead Time"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   43
      Left            =   240
      TabIndex        =   21
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Manufacturing Flow Time"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   44
      Left            =   240
      TabIndex        =   20
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(Days)"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   45
      Left            =   4200
      TabIndex        =   19
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(Days)"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   46
      Left            =   4200
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "ShopSHe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/19/06 Revised form and labels
Option Explicit
Dim bOnLoad As Byte

Dim iCurrIndex As Integer
Dim iTotalParts As Integer
Dim vPartParam(200, 6) As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetShopDefaults()
   Dim RdoShp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetLeadTimes"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         txtDefFlow = Format(!DEFFLOWTIME, "##0")
         txtDefLead = Format(!DEFLEADTIME, "##0")
         ClearResultSet RdoShp
      End With
   End If
   Set RdoShp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getshopdef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbLvl_LostFocus()
   Dim b1 As Byte
   Dim b2 As Byte
   cmbLvl = CheckLen(cmbLvl, 3)
   For b2 = 0 To 8
      If cmbLvl.List(b2) = cmbLvl Then b1 = 1
   Next
   If b1 = 0 Then
      Beep
      cmbLvl = cmbLvl.List(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4105
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdNxt_Click()
   iCurrIndex = iCurrIndex + 6
   GetNextGroup
   
End Sub

Private Sub cmdPrev_Click()
  If iCurrIndex > 0 Then
      iCurrIndex = iCurrIndex - 6
      GetNextGroup
  End If
End Sub


Private Sub cmdSel_Click()
   FillPartParams
   
End Sub

Private Sub Form_Activate()
   Dim b As Byte
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetShopDefaults
      bOnLoad = 0
   End If
   cmbLvl.AddItem "ALL"
   For b = 1 To 7
      cmbLvl.AddItem Trim(str(b))
   Next
   cmbLvl.AddItem Trim(str(b))
   cmbLvl = cmbLvl.List(0)
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim iList As Integer
   FormLoad Me
   FormatControls
   
   For iList = 0 To 5
      txtFlw(iList).Enabled = False
      txtFlw(iList).Text = ""
      txtFlw(iList).BackColor = vbButtonFace
      txtLed(iList).Enabled = False
      txtLed(iList).Text = ""
      txtLed(iList).BackColor = vbButtonFace
   Next
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set ShopSHe05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub



Private Sub txtDefFlow_LostFocus()
   On Error Resume Next
   txtDefFlow = CheckLen(txtDefFlow, 3)
   txtDefFlow = Format(Abs(Val(txtDefFlow)), "##0")
   sSql = "UPDATE Preferences SET DEFFLOWTIME=" & Val(Trim(txtDefFlow)) & " " _
          & "WHERE PreRecord=1"
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtDefLead_LostFocus()
   On Error Resume Next
   txtDefLead = CheckLen(txtDefLead, 3)
   txtDefLead = Format(Abs(Val(txtDefLead)), "##0")
   sSql = "UPDATE Preferences SET DEFLEADTIME=" & Val(Trim(txtDefLead)) & " " _
          & "WHERE PreRecord=1"
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtFlw_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtFlw_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue (KeyAscii)
   
End Sub

Private Sub txtFlw_LostFocus(Index As Integer)
   On Error Resume Next
   txtFlw(Index) = CheckLen(txtFlw(Index), 3)
   txtFlw(Index) = Format(Abs(Val(txtFlw(Index))), "##0")
   If vPartParam(Index + iCurrIndex, 4) <> Val(txtFlw(Index)) Then
      sSql = "UPDATE PartTable SET PAFLOWTIME=" & Val(Trim(txtFlw(Index))) & " " _
             & "WHERE PARTREF='" & Compress(lblPrt(Index)) & "'"
      clsADOCon.ExecuteSQL sSql
      vPartParam(Index + iCurrIndex, 4) = Val(txtFlw(Index))
   End If
   
   
End Sub


Private Sub txtLed_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtLed_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub

Private Sub txtLed_LostFocus(Index As Integer)
   On Error Resume Next
   txtLed(Index) = CheckLen(txtLed(Index), 3)
   txtLed(Index) = Format(Abs(Val(txtLed(Index))), "##0")
   If vPartParam(Index + iCurrIndex, 5) <> Val(txtLed(Index)) Then
      sSql = "UPDATE PartTable SET PALEADTIME=" & Val(Trim(txtLed(Index))) & " " _
             & "WHERE PARTREF='" & Compress(lblPrt(Index)) & "'"
      clsADOCon.ExecuteSQL sSql
      vPartParam(Index + iCurrIndex, 5) = Val(txtLed(Index))
   End If
   
End Sub



Private Sub FillPartParams()
   Dim RdoPrm As ADODB.Recordset
   Dim bLen As Byte
   Dim iRow As Integer
   Dim sPart As String
   Erase vPartParam
   If txtPrt <> "ALL" Then sPart = Compress(txtPrt)
   bLen = Len(sPart)
   If bLen = 0 Then
      bLen = 1
      txtPrt = "ALL"
   End If
   iTotalParts = -1
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PALEADTIME," _
          & "PAFLOWTIME FROM PartTable WHERE LEFT(PARTREF," _
          & bLen & ") >= '" & Left(sPart, bLen) & "' "
   If cmbLvl <> "ALL" Then sSql = sSql & "AND PALEVEL=" & cmbLvl & " "
   sSql = sSql & "ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrm, ES_FORWARD)
   If bSqlRows Then
      With RdoPrm
         Do Until .EOF
            iTotalParts = iTotalParts + 1
            iRow = iTotalParts
            If iTotalParts > 199 Then Exit Do
            vPartParam(iRow, 0) = "" & Trim(!PartRef)
            vPartParam(iRow, 1) = "" & Trim(!PartNum)
            vPartParam(iRow, 2) = "" & Trim(!PADESC)
            vPartParam(iRow, 3) = "" & Trim(!PALEVEL)
            vPartParam(iRow, 4) = "" & Trim(!PAFLOWTIME)
            vPartParam(iRow, 5) = "" & Trim(!PALEADTIME)
            .MoveNext
         Loop
         ClearResultSet RdoPrm
      End With
      iCurrIndex = 0
      If iTotalParts >= 0 Then GetNextGroup
   End If
   
   Set RdoPrm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "FillPartPara"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextGroup()
   Dim iList As Integer
   For iList = 0 To 5
      lblPrt(iList) = ""
      lblDsc(iList) = ""
      lblLvl(iList) = ""
      txtFlw(iList).Enabled = False
      txtFlw(iList).Text = ""
      txtFlw(iList).BackColor = vbButtonFace
      txtLed(iList).Enabled = False
      txtLed(iList).Text = ""
      txtLed(iList).BackColor = vbButtonFace
   Next
   
   For iList = 0 To 5
      If iList + iCurrIndex > iTotalParts Then Exit For
      lblPrt(iList) = vPartParam(iList + iCurrIndex, 1)
      lblDsc(iList) = vPartParam(iList + iCurrIndex, 2)
      lblLvl(iList) = vPartParam(iList + iCurrIndex, 3)
      txtFlw(iList).Enabled = True
      txtFlw(iList).Text = vPartParam(iList + iCurrIndex, 4)
      txtFlw(iList).BackColor = vbWindowBackground
      txtLed(iList).Enabled = True
      txtLed(iList).Text = vPartParam(iList + iCurrIndex, 5)
      txtLed(iList).BackColor = vbWindowBackground
   Next
   On Error Resume Next
   If txtFlw(0).Enabled = True Then txtFlw(0).SetFocus
   
   
End Sub

Private Sub txtPrt_LostFocus()
   If Trim(txtPrt) = "" Then txtPrt = "ALL"
   
End Sub
