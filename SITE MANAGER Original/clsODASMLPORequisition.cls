VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASLPORequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ORDate As Variant, ReqDate As Variant, PromDate As Variant

Public Sub AddNewRecord()
If EditRecord Then Exit Sub
    NewRecord = True
'    With frmODASMPurchaseOrders
    
'        Call clearPRODUCT
'
'        .cmdSAVE.Caption = "&SAVE RECORD"
'        .cmdFinish.Enabled = False
'
'        If .txtOrderNo.Text = Empty Then
'            .txtOrderNo.SetFocus
'        Else
'            .txtProductDescription.SetFocus
'        End If
'
'    End With
End Sub

Private Function GetMyCompanyCode() As String
On Error GoTo err

    Set rsFindRecord = cnCOMMON.Execute("SELECT CompanyCode FROM ParamCompanyMaster WHERE CompanyCode IS NOT NULL;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetMyCompanyCode = Empty
    ElseIf IsNull(rsFindRecord!CompanyCode) Or rsFindRecord!CompanyCode = "" Then
        GetMyCompanyCode = Empty
    Else
        GetMyCompanyCode = rsFindRecord!CompanyCode
    End If
    
    Set rsFindRecord = Nothing
    
Exit Function
err:
    ErrorMessage
End Function

Private Function CurrentStaffID() As String
On Error GoTo err

    Set rsFindRecord = cnCOMMON.Execute("SELECT StaffIDNO FROM AdminUserRegister WHERE UserName='" & Trim(CurrentUserName) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        CurrentStaffID = Empty
    ElseIf IsNull(rsFindRecord!staffidno) Or rsFindRecord!staffidno = "" Then
        CurrentStaffID = Empty
    Else
        CurrentStaffID = rsFindRecord!staffidno
    End If
    
    Set rsFindRecord = Nothing
    
Exit Function
err:
    ErrorMessage
End Function
