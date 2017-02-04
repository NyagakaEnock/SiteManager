Attribute VB_Name = "modNavigateControl"
Option Explicit

Public Navigator As Boolean, CancelNavigate As Boolean
Public Editor As Boolean, CancelEditor As Boolean

Public Sub NavigateControl()
On Error Resume Next
    If Navigator = True Then
        MsgBox "Process Terminated! This operation is not permitted when Data Entry or Edit is in progress!", vbCritical, "Process Controller"
        CancelNavigate = True
    Else
        CancelNavigate = False
       Exit Sub
    End If
End Sub

Public Sub EditControl()
On Error Resume Next
    If Editor = True Then
        MsgBox "Process Cancelled! This operation is not permitted when Data Entry or Edit is in progress!", vbCritical, "Process Controller"
        CancelEditor = True
    Else
        CancelEditor = False
       Exit Sub
    End If
End Sub

