VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPartClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub PopulatePartClassCombo(cbo As ComboBox, Optional ShowAll As Boolean)
   'populate a part class combo box
   'ShowAll = True to include an <ALL> entry

   If ShowAll Then
      sSql = "SELECT '<ALL>' AS CCCODE UNION SELECT CCCODE FROM PclsTable ORDER BY CCCODE"
   Else
      sSql = "SELECT CCCODE FROM PclsTable ORDER BY CCCODE"
   End If
   LoadComboBox cbo, -1
End Sub


