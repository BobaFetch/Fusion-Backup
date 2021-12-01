VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eMRPTypes
   'ins
   MRPTYPE_BeginningBalance = 1
   MRPTYPE_PoItem = 2
   MRPTYPE_MoNoPicklist = 3
   MRPTYPE_MoWithPicklist = 4
   MRPTYPE_PoActionItem = 5
   MRPTYPE_MoActionItem = 6
   
   'outs
   MRPTYPE_SoItem = 11
   MRPTYPE_MoPick = 12
   MRPTYPE_Explosion = 13
   MRPTYPE_UnschedPick = 14
   MRPTYPE_PhantomPick = 15
   MRPTYPE_PhantomExplosion = 16
   MRPTYPE_SafetyStock = 17
End Enum

