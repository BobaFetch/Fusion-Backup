VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaCLf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Sales Activity Standard Cost"
   ClientHeight    =   3270
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3270
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1560
      TabIndex        =   17
      Top             =   840
      Width           =   2775
   End
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6360
      TabIndex        =   10
      ToolTipText     =   "Build QuickBooks Export"
      Top             =   720
      Width           =   875
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox txtSta 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Tag             =   "4"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   4440
      Picture         =   "diaCLf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Show BOM Structure"
      Top             =   840
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox cmbprt1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "3"
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3270
      FormDesignWidth =   7365
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
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaCLf01a.frx":0342
      PictureDn       =   "diaCLf01a.frx":0488
   End
   Begin ComctlLib.ProgressBar prg1 
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory activities updated"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   2145
   End
   Begin VB.Label lblSOCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SO Items"
      Height          =   285
      Index           =   11
      Left            =   3480
      TabIndex        =   13
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label lblIACount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   5
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parts Like"
      Height          =   405
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1065
   End
End
Attribute VB_Name = "diaCLf01a"
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
' diaCLf01a - Update Sales Activity Standard Cost
'
' Notes:
'
' Created: (nth) 09/13/04
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim sMsg As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*********************************************************************************

Private Sub cmbPrt_GotFocus()
   SelectFormat Me
End Sub

Private Sub cmbPrt_LostFocus()
   cmbprt = CheckLen(cmbprt, 30)
   If Trim(cmbprt) = "" Then
      cmbprt = "ALL"
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdGo_Click()
   BackFillCost cmbprt, txtSta, txtEnd
End Sub

Private Sub cmdVew_Click()
   optVew.Value = vbChecked
   ViewParts.Show
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
      FillPartCombo cmbprt
      cmbprt = "ALL"
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   sCurrForm = Caption
   txtSta = Format(ES_SYSDATE, "mm/01/yy")
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaCLf01a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub BackFillCost( _
                         sPart As String, _
                         sStart As String, _
                         sEnd As String)
   
   'Dim rdoFil As ADODB.RecordSet
   'Dim RdoPrt As ADODB.RecordSet
   'Dim rdoCnt As ADODB.RecordSet
   Dim i As Single
   Dim K As Integer
   
   On Error GoTo DiaErr1
   MouseCursor 13
   sPart = Compress(sPart)
   If sPart = "ALL" Then sPart = ""
   
   '    sSql = "SELECT COUNT(DISTINCT INPART) FROM InvaTable WHERE INADATE >='" _
   '        & sStart & "' AND INADATE <= '" & sEnd & "'"
   '    If sPart <> "ALL" Then
   '        sSql = sSql & " AND INPART LIKE '" & sPart & "%'"
   '    End If
   '    bSqlRows = clsAdoCon.GetDataSet(sSql,rdoCnt)
   '    With rdoCnt
   '        If .Fields(0) > 0 Then
   '            I = 100 / .Fields(0)
   '            lblRec = .Fields(0)
   '            .Cancel
   '        End If
   '    End With
   '    Set rdoCnt = Nothing
   '
   '    sSql = "SELECT DISTINCT INPART,PATOTLABOR,PATOTMATL,PATOTEXP," _
   '        & "PATOTOH,PATOTHRS,PASTDCOST FROM InvaTable,PartTable WHERE INPART = " _
   '        & "PARTREF AND INADATE >='" & sStart & "' AND INADATE <= '" & sEnd & "'"
   '    If sPart <> "ALL" Then
   '        sSql = sSql & " AND PARTREF LIKE '" & sPart & "%'"
   '    End If
   '    bSqlRows = clsAdoCon.GetDataSet(sSql,rdoFil)
   '    If bSqlRows Then
   '        With rdoFil
   '            prg1.Max = 100
   '            prg1.Value = 0
   '            prg1.Visible = True
   '            lblRec.Visible = True
   '            lblCount.Visible = True
   '            lblCount = 0
   '            z1(4).Visible = True
   '            z1(11).Visible = True
   '            DoEvents
   '            On Error Resume Next
   '            clsADOCon.BeginTrans
   '            Do Until .EOF
   '                k = k + 1
   '                lblCount = k
   '                lblCount.Refresh
   '                prg1.Value = prg1.Value + I
   '                sSql = "UPDATE InvaTable SET INTOTLABOR = " _
   '                    & !PATOTLABOR & ",INTOTMATL = " & !PATOTMATL _
   '                    & ",INTOTEXP=" & !PATOTEXP & ",INTOTOH = " _
   '                    & !PATOTOH & ",INTOTHRS=" & !PATOTHRS _
   '                    & ",INAMT='" & !PASTDCOST & vbCrLf _
   '                    & " WHERE INADATE >='" & sStart _
   '                    & "' AND INADATE <= '" & sEnd _
   '                    & "' AND INPART = '" & !INPART & "'"
   '                 clsAdoCon.ExecuteSQL sSQL
   '                .MoveNext
   '            Loop
   '        End With
   
   'update inventory shipped in this date range
   
   clsADOCon.ADOErrNum = 0
   
   prg1.max = 100
   prg1.Value = 0
   prg1.Visible = True
   sSql = "update InvaTable SET INTOTLABOR=PATOTLABOR," & vbCrLf _
          & "INTOTMATL=PATOTMATL, INTOTEXP=PATOTEXP," & vbCrLf _
          & "INTOTOH=PATOTOH, INTOTHRS=PATOTHRS," & vbCrLf _
          & "INAMT=PASTDCOST" & vbCrLf _
          & "from InvaTable item join PartTable part on item.INPART = part.PARTREF" & vbCrLf _
          & "where item.INADATE >= '" & sStart & "'" & vbCrLf _
          & "and item.INADATE <= '" & sEnd & "'" & vbCrLf _
          & "and part.PARTREF like '" & sPart & "%'"
   clsADOCon.ExecuteSQL sSql
   prg1.Value = 50
   lblIACount = clsADOCon.RowsAffected
   
   'update so items shipped in this date range
   sSql = "update SoitTable set ITCOST = PASTDCOST" & vbCrLf _
          & "from SoitTable item join PartTable part on item.ITPART = part.PARTREF" & vbCrLf _
          & "where item.ITACTUAL >= '" & sStart & "' and item.ITACTUAL <= '" & sEnd & "'" & vbCrLf _
          & "and part.PARTREF like '" & sPart & "%'"
   clsADOCon.ExecuteSQL sSql
   prg1.Value = 100
   lblSOCount = clsADOCon.RowsAffected
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      sMsg = "Sales Activity Standard Cost" _
             & vbCrLf & "Successfully Updated ."
      MsgBox sMsg, vbInformation, Caption
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0

      sMsg = "Cannot Update Sales Activity" _
             & vbCrLf & "Transaction Canceled."
      MsgBox sMsg, vbExclamation, Caption
   End If
   '    Else
   '        sMsg = "No Records Found."
   '        MsgBox sMsg, vbInformation, Caption
   '    End If
   MouseCursor 0
   'lblRec.Visible = False
   prg1.Visible = False
   'lblCount.Visible = False
   lblSOCount = ""
   lblIACount = ""
   'z1(4).Visible = False
   'z1(11).Visible = False
   'Set rdoFil = Nothing
   Exit Sub
DiaErr1:
   sProcName = "backfill"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtSta_LostFocus()
   txtSta = CheckDate(txtSta)
End Sub

Private Sub txtSta_DropDown()
   ShowCalendar Me
End Sub

