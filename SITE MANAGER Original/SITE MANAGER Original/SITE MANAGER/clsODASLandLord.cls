VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASLandLord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub deleteRECORD()
'On Error GoTo Myerr

If cboPolicyNo.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
ElseIf Screen.ActiveForm.txtLandLordNo.Text = "" Then
            MsgBox "There is no current record", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            With RsCode
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearALLRECORD
                
                                
            End With
    End If
        '/* End if Msg Box
        
End If
        '/* If frmALISMBeneficiary.txt = ""
        
Exit Sub

Myerr:
    ErrorMessage

End Sub

Public Sub SearchRecord()
        loadRECORD
        LoadAccountType
End Sub

Public Sub loadRECORD()
On Error GoTo err
With frmODASPLandLord

        Set rsCONTROL = New ADODB.Recordset
        rsCONTROL.Open "SELECT * FROM ODASPAccount  WHERE AccountNo = '" & .txtLandLordNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF And rsCONTROL.BOF Then
            MsgBox "The requested record does not exist in the database.", vbOKOnly + vbExclamation, "Searching"
        Else
             .txtLandLordName.Text = rsCONTROL!CompanyName
             .txtLandLordNo.Text = rsCONTROL!AccountNo
             .txtemailAddress.Text = rsCONTROL!EmailAddress & ""
             .txtMobileNo.Text = rsCONTROL!MobileNo & ""
             .txtPhysicalAddress.Text = rsCONTROL!PhysicalAddress & ""
             .txtPostalAddress.Text = rsCONTROL!PostalAddress
             .txtTelephoneNo.Text = rsCONTROL!TelephoneNo & ""
             .cboAccountType.Text = rsCONTROL!AccountType & ""
             .cboTownCode.Text = rsCONTROL!Towncity & ""
             
             .txtContactDesignation.Text = rsCONTROL!ContactTitle & ""
             .txtContactName.Text = rsCONTROL!ContactPerson & ""

             
        
        End If

        End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadLANDLORD()
On Error GoTo err
With frmODASMSiteRegistration

    Set rsCONTROL = New ADODB.Recordset
    rsCONTROL.Open "SELECT * FROM ODASPAccount  WHERE AccountNo = '" & .txtLandLordNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsCONTROL.EOF And rsCONTROL.BOF Then
        MsgBox "The requested record does not exist in the database.", vbOKOnly + vbExclamation, "Searching"
    Else
         bDontChange = True
         .txtLandLordName.Text = rsCONTROL!CompanyName
         .txtemailAddress.Text = rsCONTROL!EmailAddress & ""
         .txtMobileNo.Text = rsCONTROL!MobileNo & ""
         .txtPostalAddress.Text = rsCONTROL!PostalAddress
         .txtTownCity.Text = rsCONTROL!Towncity & ""
                     
         Set rsCONTROL = Nothing
    
    End If

End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub saveRecord()
On Error GoTo err
With Screen.ActiveForm

        Set rsSAVE = New ADODB.Recordset
        rsSAVE.Open "SELECT * FROM ODASPAccount WHERE AccountNo = '" & .txtLandLordNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsSAVE.EOF And rsSAVE.BOF Then
                rsSAVE.AddNew
                rsSAVE!AccountNo = .txtLandLordNo.Text
                rsSAVE!CompanyName = .txtLandLordName.Text
                rsSAVE!CreatedBy = InUserName
                rsSAVE!DateCreated = Date
                rsSAVE!Status = "A"
         End If
         
         rsSAVE!EmailAddress = .txtemailAddress.Text
         rsSAVE!MobileNo = .txtMobileNo.Text
         rsSAVE!PostalAddress = .txtPostalAddress.Text
         
         If bAllowProcess Then
                 rsSAVE!AccountType = "LLORD"
                 rsSAVE!Towncity = .txtTownCity.Text
         Else
                 rsSAVE!ContactTitle = .txtContactDesignation.Text
                 rsSAVE!ContactPerson = .txtContactName.Text
                 rsSAVE!PhysicalAddress = .txtPhysicalAddress.Text
                 rsSAVE!TelephoneNo = .txtTelephoneNo.Text
                 rsSAVE!AccountType = .cboAccountType.Text
                 rsSAVE!Towncity = .cboTownCode.Text
        End If
        
        rsSAVE.Update
        rsSAVE.Requery
           
        Set rsSAVE = Nothing

                    
End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub updateRECORD()
        If ValidateData = True Then
            If NewRecord = True Then
                    generateLandLordNO
            End If
            saveRecord
            disableALLRECORD
            NewRecord = False
            getLANDLORDS
    End If
End Sub
Public Sub updateRECORDExternally()
With frmODASMSiteRegistration
        If NewRecord = True And Trim(.txtLandLordNo.Text) <= Empty Then generateLandLordNO
        saveRecord
End With
End Sub
Private Sub generateLandLordNO()
On Error GoTo err
With Screen.ActiveForm
        
        Set rsCONTROL = New Recordset
        strSQL = "Select * from ODASPLAstNumbers Where AutoLandLordNo = 'Y';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF = True Then Exit Sub
        
        .txtLandLordNo.Text = rsCONTROL!LandLordNo
        
        Select Case Len(.txtLandLordNo.Text)
                Case 1: .txtLandLordNo.Text = Trim(rsCONTROL!LandLordPrefix) + "0000" + Trim(rsCONTROL!LandLordNo)
                Case 2: .txtLandLordNo.Text = Trim(rsCONTROL!LandLordPrefix) + "000" + Trim(rsCONTROL!LandLordNo)
                Case 3: .txtLandLordNo.Text = Trim(rsCONTROL!LandLordPrefix) + "00" + Trim(rsCONTROL!LandLordNo)
                Case 4: .txtLandLordNo.Text = Trim(rsCONTROL!LandLordPrefix) + "0" + Trim(rsCONTROL!LandLordNo)
                Case 5: .txtLandLordNo.Text = Trim(rsCONTROL!LandLordPrefix) + Trim(rsCONTROL!LandLordNo)
        End Select
                                
        rsCONTROL!LandLordNo = Str(Val(rsCONTROL!LandLordNo) + 1)
        rsCONTROL.Update
        rsCONTROL.Requery
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Function ValidateData()
On Error GoTo err
    
    ValidateData = False
    
    With frmODASPLandLord
        If .txtLandLordName.Text <= "" Then
               MsgBox "The Account Name Cannot be Left Blank", vbOKOnly
                .txtLandLordName.SetFocus
                
        ElseIf .txtAccountTypeDescription.Text <= "" Then
                MsgBox "The Account Type Entered is invalid", vbOKOnly
                .cboAccountType.SetFocus
    
        ElseIf .txtPhysicalAddress.Text <= "" Then
                MsgBox "The Physical Address cannot be Left Blank", vbOKOnly
                .txtPhysicalAddress.SetFocus
        
        
        ElseIf .txtPostalAddress.Text = "" Then
                MsgBox "The Postal Address of the LandLord is Required", vbOKOnly
                .txtPostalAddress.SetFocus
                
        ElseIf .txtTelephoneNo.Text <= "" Then
                MsgBox "The Telephone Contact is needed for ease of Access", vbOKOnly
                .txtTelephoneNo.SetFocus
        
        ElseIf .txtTownDescription.Text <= "" Then
                MsgBox "The Town Code is neccessary", vbOKOnly
                .cboTownCode.SetFocus
        
        ElseIf .cboAccountType.Text <= "" Then
                MsgBox "The Account Type is Mandatory", vbOKOnly
                .cboAccountType.SetFocus
        
        ElseIf .cboTownCode.Text <= "" Then
                MsgBox "The Town Code cannot be Left Blank", vbOKOnly
                .cboTownCode.SetFocus
        Else
                ValidateData = True
        End If
        
    
    End With
Exit Function

err:
    ErrorMessage
End Function

Public Sub LoadDEFAULT()
On Error GoTo err
    
    With frmODASPLandLord
        .txtDateCreated.Text = Date
        .txtPostalAddress.Text = "P. O. Box "
        .txtPhysicalAddress.Text = "XX"
        .cboTownCode.Text = "NBI"
        .txtTelephoneNo.Text = "XX"
        .txtTownDescription.Text = "NAIROBI"
        .txtStatus.Text = "ACTIVE"
        .txtDateCreated.Text = Date
        .txtContactDepartment.Text = "XX"
        .txtContactDesignation.Text = "XX"
        .txtContactName.Text = "XX"
        .txtMobileNo.Text = "XX"
        .txtTelephoneExtention.Text = "XX"
        .txtemailAddress.Text = "XX"
    End With

Exit Sub

err:
    ErrorMessage
End Sub

