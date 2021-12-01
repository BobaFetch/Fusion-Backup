VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassWorkCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub PopulateShopCombo(cboShop As ComboBox, cboWorkCenter As ComboBox)
   cboShop.Clear
   DoEvents
   sSql = "Qry_FillShops "
   LoadComboBox cboShop
   DoEvents
End Sub

Public Sub PoulateWorkCenterCombo(cboShop As ComboBox, cboWorkCenter As ComboBox)
   cboWorkCenter.Clear
   DoEvents
   sSql = "Qry_FillWorkCenters '" & Compress(cboShop) & "'"
   LoadComboBox cboWorkCenter
   DoEvents
   If cboWorkCenter.ListCount > 0 Then
      cboWorkCenter.AddItem "ALL"
      cboWorkCenter = "ALL"
   End If
   DoEvents
End Sub

