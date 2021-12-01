VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProductCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub PopulateProductCodeCombo(cbo As ComboBox, Optional ShowAll As Boolean)
   'populate a product code combo box
   'ShowAll = True to include an <ALL> entry
   ';Dim Ado As ADODB.Recordset
   If ShowAll Then
      sSql = "SELECT '<ALL>' AS PCCODE UNION SELECT PCCODE FROM PcodTable ORDER BY PCCODE"
   Else
      sSql = "SELECT PCCODE FROM PcodTable ORDER BY PCCODE"
   End If
   LoadComboBox cbo, -1
End Sub


