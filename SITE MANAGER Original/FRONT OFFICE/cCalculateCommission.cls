VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCalculateCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public rsPOLICY, rsPAYMENTMODE, rsAGENTNO, rsSUPERVISOR, rsPREMPAY As ADODB.Recordset
Public strPOLICY, strPAYMENTMODE, strAGENTNO, strSUPERVISOR, strPREMPAY As String
