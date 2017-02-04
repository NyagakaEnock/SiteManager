Attribute VB_Name = "ModCommandButtons"
Option Explicit

Public Sub Cancelrecord()
        baddRECORD = False
        beditRECORD = False
        enableButtons
        Screen.ActiveForm.ClearControls
        disableALLRECORD
End Sub
