VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim saveDATA As Boolean
Public rsREFERENCE As ADODB.Recordset

'/********************************************
'/Methods for clearing the screens begin here'/
'/********************************************

Public Sub clearREFERENCECONTROLS()
'On Error GoTo err
        With frmODASPAccount
            .txtAccountNo.Text = ""
            .txtCompanyName.Text = ""
            .txtOtherNames.Text = ""
            .cboTitleCode.Text = ""
            .txtPostalAddress.Text = ""
            .txtPhysicalAddress.Text = ""
            .cboTownCode.Text = ""
            .txtTelephoneNo.Text = ""
            .txtBirthPlace.Text = ""
            .txtBirthDate.Text = ""
            .txtIdentityNo.Text = ""
            .cboIdType.Text = ""
            .cboAccountNo.Text = ""
            .cboOccupationCode.Text = ""
            .txtOccupationDescription.Text = ""
            .cboNationalityCode.Text = ""
            .cboGenderCode.Text = ""
            .txtAccountNo.Text = ""
    End With
    Exit Sub
err:
ErrorMessage
End Sub

Public Sub DisableREFERENCECONTROLS()
'On Error GoTo err
Dim Bval As Boolean
   Call PolControls(True)
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub enableREFERENCECONTROLS()
'On Error GoTo err
    Dim Bval As Boolean
    Call PolControls(False)
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub PolControls(Bval As Boolean)
With frmODASPAccount
    .txtAccountNo.Locked = Bval
    .txtCompanyName.Locked = Bval
    .txtOtherNames.Locked = Bval
    .cboTitleCode.Locked = Bval
    .txtPostalAddress.Locked = Bval
    .txtPhysicalAddress.Locked = Bval
    .cboTownCode.Locked = Bval
    .txtTelephoneNo.Locked = Bval
    .txtBirthPlace.Locked = Bval
    .txtBirthDate.Locked = Bval
    .txtIdentityNo.Locked = Bval
    .cboIdType.Locked = Bval
    .cboAccountNo.Locked = Bval
    .cboOccupationCode.Locked = Bval
    .txtOccupationDescription.Locked = Bval
    .cboNationalityCode.Locked = Bval
    .cboGenderCode.Locked = Bval
    .txtAccountNo.Locked = Bval
End With
End Sub

Public Sub LoadDESCRIPTION()
'On Error GoTo err
    Dim rsTCLF As ADODB.Recordset, strtclf As String
    Set rsTCLF = New Recordset
    
    rsTCLF.Open "SELECT * FROM ODASPAccount WHERE AccountNo= '" & frmODASPAccount.cboAccountNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsTCLF
        If .EOF And .BOF Then Exit Sub
        'frmODASPAccount.txtEmployerDescription.Text = !CompanyName
    End With
    
    Dim rsOCLF As ADODB.Recordset, strOCLF As String
    Set rsOCLF = New Recordset
    
    rsOCLF.Open "SELECT * FROM ALISPOccupation WHERE OccupationCode = '" & frmODASPAccount.cboOccupationCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsOCLF
        If .EOF And .BOF Then Exit Sub
        frmODASPAccount.txtOccupationDescription = !OccupationDescription
    End With
    Exit Sub
err:
    ErrorMessage
End Sub

Sub loadDEFAULTS()
'On Error GoTo err
    Dim rsDEFA As ADODB.Recordset, strDEFA, LSetupCode As String
    Set rsDEFA = New Recordset
    
    LSetupCode = 0
    
    strDEFA = "Select * from ALISPReferenceDefaults where setupCode =  '" & LSetupCode & "';"
    rsDEFA.Open strDEFA, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsDEFA
        If .EOF And .BOF Then
                    MsgBox "The Record Does Not Exist, Try Again !!", vbOKOnly
            Exit Sub
        End If
        
            frmODASPAccount.txtCompanyName = !CompanyName
            frmODASPAccount.txtOtherNames = !OtherNames
            frmODASPAccount.cboTitleCode = !titlecode
            frmODASPAccount.txtPostalAddress = !PostalAddress
            frmODASPAccount.txtPhysicalAddress = !PhysicalAddress
            frmODASPAccount.cboTownCode = !TownCode
            frmODASPAccount.txtTelephoneNo = !TelephoneNo
            frmODASPAccount.txtBirthPlace = !BirthPlace
            frmODASPAccount.txtBirthDate = !BirthDate
            frmODASPAccount.txtIdentityNo = !IdentityNo
            frmODASPAccount.cboIdType = !IDType
            frmODASPAccount.cboAccountNo = !AccountCode
            frmODASPAccount.cboOccupationCode = !OccupationCode
            frmODASPAccount.cboNationalityCode = !NationalityCode
            frmODASPAccount.cboGenderCode = !GenderCode
            frmODASPAccount.txtAccountNo = !AccountNo
   End With

rsDEFA.Close
strDEFA = ""


Exit Sub

err:
    ErrorMessage
End Sub

Sub loadDEFAULT()
'On Error GoTo err
    Dim rsDEFA As ADODB.Recordset, strDEFA, LSetupCode As String
    Set rsDEFA = New Recordset
    
    LSetupCode = 0
    
    strDEFA = "Select * from ALISPReferenceDefaults where setupCode =  '" & LSetupCode & "';"
    rsDEFA.Open strDEFA, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsDEFA
        If .EOF And .BOF Then
                    MsgBox "The Record Does Not Exist, Try Again !!", vbOKOnly
            Exit Sub
        End If
        
            frmODASPAccount.txtCompanyName = !CompanyName
            frmODASPAccount.txtOtherNames = !OtherNames
            frmODASPAccount.cboTitleCode = !titlecode
            frmODASPAccount.txtPostalAddress = !PostalAddress
            frmODASPAccount.txtPhysicalAddress = !PhysicalAddress
            frmODASPAccount.cboTownCode = !TownCode
            frmODASPAccount.txtTelephoneNo = !TelephoneNo
            frmODASPAccount.txtBirthPlace = !BirthPlace
            frmODASPAccount.txtBirthDate = !BirthDate
            frmODASPAccount.txtIdentityNo = !IdentityNo
            frmODASPAccount.cboIdType = !IDType
            frmODASPAccount.cboAccountNo = !AccountCode
            frmODASPAccount.cboOccupationCode = !OccupationCode
            frmODASPAccount.cboNationalityCode = !NationalityCode
            frmODASPAccount.cboGenderCode = !GenderCode
            frmODASPAccount.txtAccountNo = !AccountNo
   End With

rsDEFA.Close
strDEFA = ""


Exit Sub

err:
    ErrorMessage
End Sub


'/************************************************************************
'/The Methods for loading the records to the customer tree view begin here
'/************************************************************************

           
Private Sub validateREFERENCE()
'On Error GoTo err
    saveDATA = False

    If frmODASPAccount.txtAccountNo.Text = "" Then
            MsgBox "AccountNo is Required"
            frmODASPAccount.txtAccountNo.SetFocus
    
    ElseIf frmODASPAccount.txtCompanyName.Text = "" Then
            MsgBox "The CompanyName is  required"
            frmODASPAccount.txtCompanyName.SetFocus
    
    ElseIf frmODASPAccount.txtOtherNames.Text = "" Then
            MsgBox "The OtherNames is  required"
            frmODASPAccount.txtOtherNames.SetFocus
  
    ElseIf frmODASPAccount.cboTitleCode.Text = "" Then
            MsgBox "The Title Code is  required"
            frmODASPAccount.cboTitleCode.SetFocus
 
    ElseIf frmODASPAccount.txtPostalAddress.Text = "" Then
            MsgBox "The Postal Address is  required"
            frmODASPAccount.txtPostalAddress.SetFocus
            
     ElseIf frmODASPAccount.txtBirthDate.Text = "" Then
            MsgBox "The Date of Birth is  required"
            frmODASPAccount.txtBirthDate.SetFocus
    
    ElseIf frmODASPAccount.cboTownCode.Text = "" Then
            MsgBox "The Town is  required"
            frmODASPAccount.cboTownCode.SetFocus
    
    ElseIf frmODASPAccount.cboOccupationCode.Text = "" Then
            MsgBox "The Occupation code is  required"
            frmODASPAccount.cboOccupationCode.SetFocus
    
    ElseIf frmODASPAccount.cboNationalityCode.Text = "" Then
            MsgBox "The Nationality code is  required"
            frmODASPAccount.cboNationalityCode.SetFocus
    Else
            saveDATA = True
    End If
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub saveREFERENCE()
'On Error GoTo err
        Dim RsCode As ADODB.Recordset, strcode As String
        Set RsCode = New Recordset
      
        strcode = "SELECT * FROM ODASPAccount;"
        RsCode.Open strcode, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With RsCode
            .AddNew
            !AccountNo = frmODASPAccount.txtAccountNo
            !CompanyName = frmODASPAccount.txtCompanyName
            !OtherNames = frmODASPAccount.txtOtherNames
            !titlecode = frmODASPAccount.cboTitleCode
            !PostalAddress = frmODASPAccount.txtPostalAddress
            !PhysicalAddress = frmODASPAccount.txtPhysicalAddress & ""
            !TownCode = frmODASPAccount.cboTownCode
            !TelephoneNo = frmODASPAccount.txtTelephoneNo & ""
            !BirthPlace = frmODASPAccount.txtBirthPlace
            !BirthDate = frmODASPAccount.txtBirthDate
            !IdentityNo = frmODASPAccount.txtIdentityNo & ""
            !IDType = frmODASPAccount.cboIdType & ""
            !AccountCode = frmODASPAccount.cboAccountNo
            !OccupationCode = frmODASPAccount.cboOccupationCode
            !NationalityCode = frmODASPAccount.cboNationalityCode
            !GenderCode = frmODASPAccount.cboGenderCode
            !AccountNo = frmODASPAccount.txtAccountNo
            '!myphoto = frmODASPAccount.CommonDialog1.FileName
            .Update
            .Requery
        End With
Exit Sub
err:
    ErrorMessage
End Sub


'/**********************************************************
'/Methods for loading the records to the listview begin here
'/**********************************************************

Private Sub Class_Initialize()
'On Error GoTo err
    'Open the connection
    Set cnCOMMON = New ADODB.Connection
    cnCOMMON.ConnectionString = "PROVIDER=MSDASQL;DSN=DSNALIS;UID=;PWD=;"
    cnCOMMON.Open
    'Begin a transaction
    cnCOMMON.BeginTrans
       
    Set rsREFERENCE = New ADODB.Recordset
    rsREF.Open "SELECT * FROM ODASPAccount;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    'Call EnableCBPolicyControls
    
    'Commit the transaction
    cnCOMMON.CommitTrans

Exit Sub
err:
'Incase of an error the connection object shoud call its rollback object
'if an error occured, roll back the changes
cnCOMMON.RollbackTrans
MsgBox "An error occured changing the records.", vbExclamation
ErrorMessage
End Sub

Private Sub Class_Terminate()
'On Error GoTo err
    'Close the instance of a recordset
    rsREFERENCE.Close
    
    'Destroy the instance of the recordset
    Set rsREFERENCE = Nothing
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub loadGRID()
'On Error GoTo err:

If frmODASPAccount.txtCompanyName.Text <= "" Then Exit Sub

Dim rsGRID As ADODB.Recordset, StrGRID As String
Set rsGRID = New Recordset

    rsGRID.Open "SELECT AccountNo, CompanyName, OtherNames, IDType, IdentityNo, BirthDate FROM ODASPAccount WHERE CompanyName =  '" & frmODASPAccount.txtCompanyName.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
      
    With rsGRID
        If .EOF Or .BOF Then Exit Sub
        'Set frmODASPAccount.ReferenceGRID.DataSource = rsGRID
        MsgBox "Kindly Confirm whether this is the same Reference", vbOKCancel
    End With

Exit Sub
  
err:
    ErrorMessage
End Sub

Public Sub LoadGridBirthDate()
'On Error GoTo err:

    If frmODASPAccount.txtBirthDate.Text <= "" Then
            MsgBox "The BirthDate cannot be Left Blank", vbOKOnly
            Exit Sub
            frmODASPAccount.txtBirthDate.SetFocus
    ElseIf frmODASPAccount.txtCompanyName.Text <= "" Then
            MsgBox "The CompanyName MUST be Entered Prior to Comming to This field", vbOKOnly
            Exit Sub
            frmODASPAccount.txtCompanyName.SetFocus
    ElseIf frmODASPAccount.txtOtherNames.Text <= "" Then
           MsgBox "The otherNamesMUST be Entered Prior to Comming to This field", vbOKOnly
           frmODASPAccount.txtOtherNames.SetFocus
            Exit Sub
    Else
                    
    End If

Exit Sub

err:
    ErrorMessage
End Sub


