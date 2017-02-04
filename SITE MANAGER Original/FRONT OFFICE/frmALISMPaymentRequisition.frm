VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmALISMPaymentRequisition 
   Appearance      =   0  'Flat
   BackColor       =   &H80000016&
   Caption         =   "Payment Requisition"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9975
   Icon            =   "frmALISMPaymentRequisition.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   3015
      Left            =   6600
      TabIndex        =   21
      Top             =   2640
      Width           =   3255
      Begin VB.TextBox txtVoucherItemNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtJobBriefNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         TabIndex        =   35
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtTotalVoucherAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         TabIndex        =   33
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtVoucherAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1560
         TabIndex        =   26
         Top             =   1740
         Width           =   1455
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         TabIndex        =   24
         Top             =   1290
         Width           =   1455
      End
      Begin VB.TextBox cboDocumentNo 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Item No"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Job Brief"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2190
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Total Amount "
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2700
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Voucher Amount "
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1770
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Invoice Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Document No"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   900
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   6375
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   9735
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFC0C0&
         Height          =   600
         Left            =   1320
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   240
         Width           =   8175
      End
      Begin VB.Label Label15 
         Caption         =   "Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   413
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9735
      Begin VB.TextBox txtItems 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtRequisitionDate 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4080
         TabIndex        =   32
         Top             =   405
         Width           =   1575
      End
      Begin VB.TextBox cboPaymentCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4080
         TabIndex        =   31
         Top             =   750
         Width           =   1575
      End
      Begin VB.TextBox txtPaymentDescription 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2520
         TabIndex        =   30
         Top             =   1095
         Width           =   3135
      End
      Begin VB.TextBox txtPayeeDetails 
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1417
         Width           =   4335
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6840
         TabIndex        =   20
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1095
         Width           =   1215
      End
      Begin VB.TextBox txtPaymentCodeDescription 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   5640
         TabIndex        =   16
         Top             =   750
         Width           =   3735
      End
      Begin VB.TextBox txtCostCenter 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Top             =   750
         Width           =   1215
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox txtReference 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1095
         Width           =   1215
      End
      Begin VB.TextBox txtVoucherNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Items"
         Height          =   255
         Left            =   8160
         TabIndex        =   43
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Product Code"
         Height          =   255
         Left            =   5760
         TabIndex        =   28
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Account No"
         Height          =   255
         Left            =   5760
         TabIndex        =   19
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Payee"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Cost Center"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Status"
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Label1"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Reference"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   " Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   435
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Code"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Voucher No"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   435
         Width           =   1095
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New &Record "
            Key             =   "N"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit/Change "
            Key             =   "E"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search/Find "
            Key             =   "S"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Key             =   "F"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10560
         Top             =   -120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   1800
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmALISMPaymentRequisition.frx":0442
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmALISMPaymentRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLOANVALUE As cLoanApprovers
Dim rsLOADGRID As clsALISGRID
Dim rsPAYREQ As clsPaymentRequisition
Dim bCostCenter, bPaymentCode, bDocumentCode As Boolean
Public rsClaimApproval As clsALISApproval
Dim rsPAYMENT As clsALISPaymentRequisition

Private Sub cboPaymentCode_GotFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.PaymentCodeGotFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub cboPaymentCode_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub cboPaymentCode_LostFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.PaymentCodeLostFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub cboDocumentNo_GotFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.DocumentNoGotFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub cboDocumentNo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboDocumentNo_LostFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.DocumentNoLostFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub loadPaymentGRID()
'On Error GoTo err

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset

    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub
    rsGRID.Open "SELECT ALISPClaimType.ClaimTypeDescription, ALISMClaim.Amount FROM ALISMClaim, ALISPClaimType WHERE ALISMClaim.claimNo =  '" & frmODASMVoucher.cboDocumentNo & "' and ALISMClaim.type = 'A' and ALISMClaim.ClaimType = ALISPClaimType.ClaimType;", cnCOMMON, adOpenKeyset, adLockOptimistic
 
    Set Screen.ActiveForm.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub LoadDeductionGRID()
'On Error GoTo err
    Dim rsGRID1 As ADODB.Recordset
    Set rsGRID1 = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID1.Open "SELECT ALISPClaimType.ClaimTypeDescription, ALISMClaim.Amount FROM ALISMClaim, ALISPClaimType WHERE ALISMClaim.claimNo =  '" & frmODASMVoucher.cboDocumentNo & "' and ALISMClaim.type = 'D' and ALISMClaim.ClaimType = ALISPClaimType.ClaimType;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.DataSource = rsGRID1
    
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub LoadRejectedGRID()
'On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ALISMVoucher WHERE Status =  'REQ-PREP' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.RejectionGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadPendingGRID()
'On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ALISMVoucher WHERE Status =  'REQ-PREP' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.PendingGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadApprovedGRID()
'On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ALISMVoucher WHERE Status =  'REQ APPROVAL' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.ApprovedGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadAuthorizedGRID()
'On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ALISMVoucher WHERE Status =  'REQ AUTHORIZATION' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.AuthorizedGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub


Private Sub loadTOTALS()
'On Error GoTo err
    Dim rsTOTALS As ADODB.Recordset
    Set rsTOTALS = New ADODB.Recordset
    
    rsTOTALS.Open "SELECT * FROM ALISMClaimTotal WHERE claimNo =  '" & frmODASMVoucher.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsTOTALS
            If .BOF Or .EOF = True Then Exit Sub
            frmODASMVoucher.txtTotalPayments.Text = !proceeds
            frmODASMVoucher.txtTotalDeductions.Text = !deductions

    End With

rsTOTALS.Close

Exit Sub
err:
    ErrorMessage
End Sub


Private Sub loadGRID()
    loadPaymentGRID
    LoadDeductionGRID
    loadPendingGRID
    LoadRejectedGRID
    loadApprovedGRID
    loadAuthorizedGRID
    loadTOTALS
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        
        With frmODASMVoucher
        Set rsPAYREQ = New clsPaymentRequisition
        
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                            If editRECORD Then Exit Sub
                            NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                            rsPAYREQ.clearRECORD
                            rsPAYREQ.enableRECORD
                    Case "&Save Record "
                    
                            bsaveRECORD = False
                            rsPAYREQ.updateRECORD
                                    
                            If bsaveRECORD = True Then
                                        bsaveRECORD = False
                                        .Toolbar1.Buttons(2).Caption = "New &Record "
                                        .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                        .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                            End If
                    
                    Case "&NEXT ITEM "
                            
                            .Toolbar1.Buttons(1).Caption = "&Save Record"
                            rsPAYREQ.enableRECORD
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                
                Case "&Save Record "

                        bsaveRECORD = False
                        rsPAYREQ.validateRECORD
                        
                        If bsaveRECORD = True Then
                                rsPAYREQ.updateRECORD
                                If bsaveRECORD = False Then
                                          .Toolbar1.Buttons(2).Caption = "New &Record "
                                          .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                          .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                                End If
                        End If
                
                Case "&NEXT ITEM "
                            .Toolbar1.Buttons(3).Caption = "&Save Record "
                            rsPAYREQ.enableRECORD
                            rsPAYREQ.clearRECORD
                Case Else
            End Select
        
        Case "S"
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD

        Case "R"
'            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
'                .Toolbar1.Buttons(2).Caption = "New &Record "
'                .Toolbar1.Buttons(2).Image = 2
'                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
'                .Toolbar1.Buttons(3).Image = 5
'                NewRecord = False: editRECORD = False: clearALLRECORD
        Case "P"
                If frmODASMVoucher.txtVoucherNo.Text >= "" Then
                        Load frmPayRequisition
                        frmPayRequisition.Show 1, Me
                End If
            
                Load frmPayRequisitionListing
                frmPayRequisitionListing.Show 1, Me

        Case "F"
     
     
        Case Else
            Exit Sub
        End Select
        
        Set rsPAYREQ = Nothing
        
End With
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub txtCostcenter_GotFocus()
    bCostCenter = True
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.selectCostCenterGotFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub txtCostcenter_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub txtCostcenter_LostFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.selectCostCenterLostFocus
    Set rsPAYREQ = Nothing
    bCostCenter = False
End Sub

Private Sub loadPaymentDescription()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.loadPaymentDescription
    Set rsPAYREQ = Nothing
End Sub

Private Sub cmdApprove_Click()
        bapproveREQUISITION = True
        bAuthorizeREQUISITION = False
        Set rsClaimApproval = New clsALISApproval
        rsClaimApproval.checkAPPROVEDDISCHARGE
        If bapproveREQUISITION = False Then Exit Sub
        rsClaimApproval.approveCLAIM
        rsClaimApproval.loadAPPROVALDETAILS
        rsClaimApproval.switchCOMMANDBUTTONS
        Set rsClaimApproval = Nothing
        bapproveREQUISITION = False
End Sub


Private Sub checkAPPROVALSTATUS()
'On Error GoTo err
                
                Dim rsAUTHORIZATION As ADODB.Recordset, strAuthorization As String
                Set rsAUTHORIZATION = New Recordset
                    
                strAuthorization = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentApproval = '1' ;"
                rsAUTHORIZATION.Open strAuthorization, cnCOMMON, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
                        Set rsAPPROVED = New Recordset
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmODASMVoucher.txtVoucherNo & "' and operationType = '" & !OperationType & "' ;"
                        rsAPPROVED.Open strAPPROVED, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                        With rsAPPROVED
                                If .BOF Or .EOF Then
                                        MsgBox "Authorization can only take place immediately after Approval", vbOKOnly
                                                bExitSub = True
                                                Exit Sub
                                ElseIf !Accept = "N" Then
                                        MsgBox "Cannot Approve payment That has been Rejected", vbOKOnly
                                                bExitSub = True
                                                Exit Sub
                                End If
                        
                        End With

            End With
                    '/ Authorization
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub checkAUTHORIZATIONSTATUS()
'On Error GoTo err
                
                Dim rsAUTHORIZATION As ADODB.Recordset, strAuthorization As String
                Set rsAUTHORIZATION = New Recordset
                    
                strAuthorization = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentAuthorization = '1' ;"
                rsAUTHORIZATION.Open strAuthorization, cnCOMMON, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
                        Set rsAPPROVED = New Recordset
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmODASMVoucher.txtVoucherNo & "' and operationType = '" & !OperationType & "' ;"
                        rsAPPROVED.Open strAPPROVED, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                        With rsAPPROVED
                                If .BOF Or .EOF Then
                                        Exit Sub
                                Else
                                        MsgBox "This Record has Already been Authorized", vbOKOnly
                                                bExitSub = True
                                End If
                        
                        End With

                        Dim rsCHKAPPROVAL As ADODB.Recordset, strCHKAPPROVAL As String
                        Set rsCHKAPPROVAL = New Recordset
                            
                        strCHKAPPROVAL = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentApproval = '1' ;"
                        rsCHKAPPROVAL.Open strCHKAPPROVAL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        With rsCHKAPPROVAL
                                If .BOF Or .EOF Then Exit Sub
                                
                                    Dim rsAPPROVE As ADODB.Recordset, strAPPROVE As String
                                    Set rsAPPROVE = New Recordset
                                        
                                    strAPPROVE = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmODASMVoucher.txtVoucherNo & "' and operationType = '" & !OperationType & "' ;"
                                    rsAPPROVE.Open strAPPROVE, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                                    With rsAPPROVE
                                            If .BOF Or .EOF Then
                                                    Exit Sub
                                            Else
                                                    MsgBox "This Record has Already been Approved", vbOKOnly
                                                            bExitSub = True
                                            End If
                                    
                                    End With
                        
                        End With

            End With
                    '/ Authorization
Exit Sub

err:
    ErrorMessage
End Sub






Private Sub cmdSearch_Click()
    bapproveREQUISITION = True
    
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.searchRECORD
    rsPAYREQ.loadPaymentDescription
    Set rsPAYREQ = Nothing
    
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.loadAPPROVALDETAILS
    Set rsClaimApproval = Nothing
    bapproveREQUISITION = False

End Sub

Private Sub cmdSearchClaimNo_Click()
    bapproveREQUISITION = True
    bUseClaimNo = True
    
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.searchRECORD
    Set rsPAYREQ = Nothing
    
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.switchCOMMANDBUTTONS
    rsClaimApproval.loadAPPROVALDETAILS
    Set rsClaimApproval = Nothing
    bUseClaimNo = False
    bapproveREQUISITION = False
End Sub





Private Sub cmdCancel_Click()
        Set rsPAYREQ = New clsPaymentRequisition
        rsPAYREQ.Cancelrecord
        Set rsPAYREQ = Nothing
End Sub

Private Sub cmdDelete_Click()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.deleteRECORD
    Set rsPAYREQ = Nothing
    
    
End Sub

Private Sub cmdEdit_Click()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.beditRECORD
    Set rsPAYREQ = Nothing
    
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.switchCOMMANDBUTTONS
    rsClaimApproval.loadAPPROVALDETAILS
    Set rsClaimApproval = Nothing

End Sub


Private Sub cmdUpdate_Click()
        Set rsPAYREQ = New clsPaymentRequisition
        rsPAYREQ.updateRECORD
        Set rsPAYREQ = Nothing
End Sub

Private Sub ClearControls()
'On Error GoTo err

        With frmODASMVoucher
                .cboDocumentNo.Text = ""
                .txtReference.Text = ""
        End With

Exit Sub

err:
    ErrorMessage
End Sub

