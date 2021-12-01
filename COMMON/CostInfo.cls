VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CostInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private part As String
Private xAmt As Currency
Private xMatl As Currency
Private xLabor As Currency
Private xExp As Currency
Private xOH As Currency
Private xHrs As Currency

Public Property Get PartRef() As String
   PartRef = part
End Property

Public Property Get Amt() As Currency
   Amt = xAmt
End Property

Public Property Get Matl() As Currency
   Matl = xMatl
End Property

Public Property Get Labor() As Currency
   Labor = xLabor
End Property

Public Property Get Exp() As Currency
   Exp = xExp
End Property

Public Property Get OH() As Currency
   OH = xOH
End Property

Public Property Get Hrs() As Currency
   Hrs = xHrs
End Property

Public Property Let PartRef(NewValue As String)
   Dim rdo As ADODB.Recordset
   part = Trim(NewValue)
   
   Dim sql As String
   sSql = "select PAUSEACTUALCOST, PALOTTRACK, PASTDCOST, PATOTMATL, " & vbCrLf _
          & "PATOTLABOR, PATOTEXP, PATOTOH, PATOTHRS" & vbCrLf _
          & "from PartTable where PARTREF = '" & part & "'"
   
   'get standard costing information from part record
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      With rdo
         If !PAUSEACTUALCOST = 0 Or !PALOTTRACK = 0 Then
            xAmt = !PASTDCOST
            xMatl = !PATOTMATL
            xLabor = !PATOTLABOR
            xExp = !PATOTEXP
            xOH = !PATOTOH
            xHrs = !PATOTHRS
         End If
         Set rdo = Nothing
         Exit Property
      End With
   End If
   Set rdo = Nothing
End Property
