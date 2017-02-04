VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASPOperationType 
   Caption         =   "Operation Type"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10425
   Icon            =   "frmODASPOperationType.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPOperationType.frx":0442
   ScaleHeight     =   6765
   ScaleWidth      =   10425
   Begin VB.Frame Frame12 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtOperationType 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtOperationDescription 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   6135
      End
      Begin VB.Frame Frame4 
         Caption         =   "Operations"
         Height          =   2775
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   9975
         Begin VB.OptionButton optsendNoticeAPPROVAL 
            Height          =   255
            Left            =   7920
            TabIndex        =   47
            Top             =   480
            Width           =   255
         End
         Begin VB.OptionButton optsendnoticeAUTHORIZATION 
            Height          =   255
            Left            =   8880
            TabIndex        =   46
            Top             =   480
            Width           =   255
         End
         Begin VB.OptionButton optSendNoticePREPARATION 
            Height          =   255
            Left            =   6600
            TabIndex        =   45
            Top             =   480
            Width           =   255
         End
         Begin VB.OptionButton optreceivenoticePREPARATION 
            Height          =   255
            Left            =   6600
            TabIndex        =   44
            Top             =   840
            Width           =   255
         End
         Begin VB.OptionButton optreceivenoticeAUTHORIZATION 
            Height          =   255
            Left            =   8880
            TabIndex        =   43
            Top             =   840
            Width           =   255
         End
         Begin VB.OptionButton optreceivenoticeAPPROVAL 
            Height          =   255
            Left            =   7920
            TabIndex        =   42
            Top             =   840
            Width           =   255
         End
         Begin VB.OptionButton optSiteApproval 
            Height          =   255
            Left            =   3000
            TabIndex        =   40
            Top             =   2400
            Width           =   255
         End
         Begin VB.OptionButton optSiteAuthorization 
            Height          =   255
            Left            =   4080
            TabIndex        =   39
            Top             =   2400
            Width           =   255
         End
         Begin VB.OptionButton optSitePreparation 
            Height          =   255
            Left            =   1920
            TabIndex        =   38
            Top             =   2400
            Width           =   255
         End
         Begin VB.OptionButton optPurchaseOrderPREPARATION 
            Height          =   255
            Left            =   1920
            TabIndex        =   36
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton optPurchaseOrderAUTHORIZATION 
            Height          =   255
            Left            =   4080
            TabIndex        =   35
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton optPurchaseOrderAPPROVAL 
            Height          =   255
            Left            =   3000
            TabIndex        =   34
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton optJobCardApproval 
            Height          =   255
            Left            =   3000
            TabIndex        =   25
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton optJobCardAuthorization 
            Height          =   255
            Left            =   4080
            TabIndex        =   24
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton optJobCardPreparation 
            Height          =   255
            Left            =   1920
            TabIndex        =   23
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton optJobBriefApproval 
            Height          =   255
            Left            =   3000
            TabIndex        =   22
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton optJobBriefAuthorization 
            Height          =   255
            Left            =   4080
            TabIndex        =   21
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton optJobBriefPreparation 
            Height          =   255
            Left            =   1920
            TabIndex        =   20
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton optQuotationPreparation 
            Height          =   255
            Left            =   1920
            TabIndex        =   19
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optQuotationAuthorization 
            Height          =   255
            Left            =   4080
            TabIndex        =   18
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optQuotationApproval 
            Height          =   255
            Left            =   3000
            TabIndex        =   17
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optInvoiceApproval 
            Height          =   255
            Left            =   3000
            TabIndex        =   16
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton optInvoiceAuthorization 
            Height          =   255
            Left            =   4080
            TabIndex        =   15
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton optInvoicePreparation 
            Height          =   255
            Left            =   1920
            TabIndex        =   14
            Top             =   1680
            Width           =   255
         End
         Begin VB.Line Line1 
            X1              =   5040
            X2              =   5040
            Y1              =   120
            Y2              =   2760
         End
         Begin VB.Label Label14 
            Caption         =   "Authorization"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8640
            TabIndex        =   52
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Approval"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7560
            TabIndex        =   51
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Preparation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   50
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Receive Notices"
            Height          =   255
            Left            =   5160
            TabIndex        =   49
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Send Notices"
            Height          =   255
            Left            =   5160
            TabIndex        =   48
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Site"
            Height          =   255
            Left            =   480
            TabIndex        =   41
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Purchase Order"
            Height          =   255
            Left            =   480
            TabIndex        =   37
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Job Card"
            Height          =   255
            Left            =   480
            TabIndex        =   32
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Quotation"
            Height          =   255
            Left            =   480
            TabIndex        =   31
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Job Brief"
            Height          =   255
            Left            =   480
            TabIndex        =   30
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Preparation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Approval"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Authorization"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Invoice"
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "options Available"
         Height          =   3015
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   8775
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4683
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   9000
         TabIndex        =   4
         Top             =   3480
         Width           =   1095
         Begin VB.CommandButton cmdAddNew 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPOperationType.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPOperationType.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   990
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPOperationType.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1365
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPOperationType.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1740
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPOperationType.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2115
            Width           =   855
         End
         Begin VB.CommandButton cmdPrint 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPOperationType.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2520
            Width           =   855
         End
      End
      Begin VB.Label lblRelationshipCode 
         Caption         =   "Operation Type"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   435
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmODASPOperationType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub loadRECORD()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPOperationType Where OperationType = '" & frmODASPOperationType.txtOperationType.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
        
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPOperationType.txtOperationType = !OperationType
        
        frmODASPOperationType.txtOperationDescription = !Description
        
        If !ReceiveNoticeApproval = True Then
                frmODASPOperationType.optreceivenoticeAPPROVAL.Value = True
        Else: frmODASPOperationType.optreceivenoticeAPPROVAL.Value = False
        End If
        
        If !ReceiveNoticePreparation = True Then
                frmODASPOperationType.optreceivenoticePREPARATION.Value = True
        Else: frmODASPOperationType.optreceivenoticePREPARATION.Value = False
        End If
        
        If !ReceiveNoticeAuthorization = True Then
                frmODASPOperationType.optreceivenoticeAUTHORIZATION.Value = True
        Else: frmODASPOperationType.optreceivenoticeAUTHORIZATION.Value = False
        End If

        If !SendNoticeApproval = True Then
                frmODASPOperationType.optsendNoticeAPPROVAL.Value = True
        Else: frmODASPOperationType.optsendNoticeAPPROVAL.Value = False
        End If
        
        If !SendNoticePreparation = True Then
                frmODASPOperationType.optSendNoticePREPARATION.Value = True
        Else: frmODASPOperationType.optSendNoticePREPARATION.Value = False
        End If
        
        If !SendNoticeAuthorization = True Then
                frmODASPOperationType.optsendnoticeAUTHORIZATION.Value = True
        Else: frmODASPOperationType.optsendnoticeAUTHORIZATION.Value = False
        End If

        If !PurchaseOrderApproval = True Then
                frmODASPOperationType.optPurchaseOrderAPPROVAL.Value = True
        Else: frmODASPOperationType.optPurchaseOrderAPPROVAL.Value = False
        End If
        
        If !PurchaseOrderPreparation = True Then
                frmODASPOperationType.optPurchaseOrderPREPARATION.Value = True
        Else: frmODASPOperationType.optPurchaseOrderPREPARATION.Value = False
        End If
        
        If !PurchaseOrderAuthorization = True Then
                frmODASPOperationType.optPurchaseOrderAUTHORIZATION.Value = True
        Else: frmODASPOperationType.optPurchaseOrderAUTHORIZATION.Value = False
        End If
                    
        If !QuotationApproval = True Then
                frmODASPOperationType.optQuotationApproval.Value = True
        Else: frmODASPOperationType.optQuotationApproval.Value = False
        End If
        
        If !QuotationPreparation = True Then
                frmODASPOperationType.optQuotationPreparation.Value = True
        Else: frmODASPOperationType.optQuotationPreparation.Value = False
        End If
        
        If !QuotationAuthorization = True Then
                frmODASPOperationType.optQuotationAuthorization.Value = True
        Else: frmODASPOperationType.optQuotationAuthorization.Value = False
        End If
    
        
        If !JobCardApproval = True Then
                frmODASPOperationType.optJobCardApproval.Value = True
        Else: frmODASPOperationType.optJobCardApproval.Value = False
        End If
        
        If !JobCardPreparation = True Then
                frmODASPOperationType.optJobCardPreparation.Value = True
        Else: frmODASPOperationType.optJobCardPreparation.Value = False
        End If
        
        If !JobCardAuthorization = True Then
                frmODASPOperationType.optJobCardAuthorization.Value = True
        Else: frmODASPOperationType.optJobCardAuthorization.Value = False
        End If
    
        
        
        If !JobBriefApproval = True Then
                frmODASPOperationType.optJobBriefApproval.Value = True
        Else: frmODASPOperationType.optJobBriefApproval.Value = False
        End If
        
        If !JobBriefPreparation = True Then
                frmODASPOperationType.optJobBriefPreparation.Value = True
        Else: frmODASPOperationType.optJobBriefPreparation.Value = False
        End If
        
        If !JobBriefAuthorization = True Then
                frmODASPOperationType.optJobBriefAuthorization.Value = True
        Else: frmODASPOperationType.optJobBriefAuthorization.Value = False
        End If
           
           
        If !InvoiceApproval = True Then
                frmODASPOperationType.optInvoiceApproval.Value = True
        Else: frmODASPOperationType.optInvoiceApproval.Value = False
        End If
        
        If !InvoicePreparation = True Then
                frmODASPOperationType.optInvoicePreparation.Value = True
        Else: frmODASPOperationType.optInvoicePreparation.Value = False
        End If
        
        If !InvoiceAuthorization = True Then
                frmODASPOperationType.optInvoiceAuthorization.Value = True
        Else: frmODASPOperationType.optInvoiceAuthorization.Value = False
        End If
        
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdAddNew_Click()
        baddRECORD = True
        clearALLRECORD
        enableALLRECORD
        'disableButtons
End Sub


Private Sub cmdCancel_Click()
'        enableButtons
        clearALLRECORD
        disableALLRECORD
        baddRECORD = False
End Sub


Private Sub cmdDelete_Click()
On Error GoTo err

If txtOperationType.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            Set rsCONTROL = New ADODB.Recordset
    
            strSQL = "Select * from ODASPOperationType Where OperationType = '" & frmODASPOperationType.txtOperationType.Text & "'"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            
            With rsCONTROL
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearALLRECORD
                getALLOPERATIONS
            End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cmdEdit_Click()
        editMYRECORD
End Sub
Public Sub GenerateOperationType()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoOperationType = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtOperationType.Text = !OperationType & ""
                
                Select Case Len(Trim(frmODASPOperationType.txtOperationType))
                        Case 1: frmODASPOperationType.txtOperationType.Text = Trim(!OperationTypePrefix) + "00" + Trim(frmODASPOperationType.txtOperationType)
                        Case 2: frmODASPOperationType.txtOperationType.Text = Trim(!OperationTypePrefix) + "0" + Trim(frmODASPOperationType.txtOperationType)
                        Case 3: frmODASPOperationType.txtOperationType.Text = Trim(!OperationTypePrefix) + Trim(frmODASPOperationType.txtOperationType)
                End Select
                
                !OperationType = !OperationType + 1
                .Update
                .Requery

        End With
        
rsCONTROL.Close

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ValidateRECORD()
On Error GoTo err

        bSaveRECORD = False
        
        If Screen.ActiveForm.txtOperationDescription.Text <= "" Then
                MsgBox "The Description of the Operation cannot be Left Blank"
                txtOperationDescription.SetFocus
        Else
                bSaveRECORD = True
        End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveRECORD()
On Error GoTo err
    Set rsSAVE = New ADODB.Recordset
    
    strSQL = "Select * from ODASPOperationType Where OperationType = '" & frmODASPOperationType.txtOperationType.Text & "'"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
   With rsSAVE
        If .BOF Or .EOF Then
                .AddNew
                GenerateOperationType
                !OperationType = frmODASPOperationType.txtOperationType
                !PreparedBy = CurrentUserName
                !DatePrepared = Date
        End If
        
        !Description = frmODASPOperationType.txtOperationDescription
        
        If frmODASPOperationType.optreceivenoticePREPARATION = True Then
                !ReceiveNoticePreparation = 1
            Else: !ReceiveNoticePreparation = 0
        End If
        
        If frmODASPOperationType.optreceivenoticeAPPROVAL = True Then
                !ReceiveNoticeApproval = 1
        Else: !ReceiveNoticeApproval = 0
        End If
        
        If frmODASPOperationType.optreceivenoticeAUTHORIZATION = True Then
                !ReceiveNoticeAuthorization = 1
        Else: !ReceiveNoticeAuthorization = 0
        End If

        If frmODASPOperationType.optSendNoticePREPARATION = True Then
                !SendNoticePreparation = 1
            Else: !SendNoticePreparation = 0
        End If
        
        If frmODASPOperationType.optsendNoticeAPPROVAL = True Then
                !SendNoticeApproval = 1
        Else: !SendNoticeApproval = 0
        End If
        
        If frmODASPOperationType.optsendnoticeAUTHORIZATION = True Then
                !SendNoticeAuthorization = 1
        Else: !SendNoticeAuthorization = 0
        End If

        If frmODASPOperationType.optSitePreparation = True Then
                !sitePreparation = 1
            Else: !sitePreparation = 0
        End If
        
        If frmODASPOperationType.optSiteApproval = True Then
                !siteApproval = 1
        Else: !siteApproval = 0
        End If
        
        If frmODASPOperationType.optSiteAuthorization = True Then
                !siteAuthorization = 1
        Else: !siteAuthorization = 0
        End If

        If frmODASPOperationType.optPurchaseOrderPREPARATION = True Then
                !PurchaseOrderPreparation = 1
            Else: !PurchaseOrderPreparation = 0
        End If
        
        If frmODASPOperationType.optPurchaseOrderAPPROVAL = True Then
                !PurchaseOrderApproval = 1
        Else: !PurchaseOrderApproval = 0
        End If
        
        If frmODASPOperationType.optPurchaseOrderAUTHORIZATION = True Then
                !PurchaseOrderAuthorization = 1
        Else: !PurchaseOrderAuthorization = 0
        End If

        If frmODASPOperationType.optQuotationPreparation = True Then
                !QuotationPreparation = 1
            Else: !QuotationPreparation = 0
        End If
        
        If frmODASPOperationType.optQuotationApproval = True Then
                !QuotationApproval = 1
        Else: !QuotationApproval = 0
        End If
        
        If frmODASPOperationType.optQuotationAuthorization = True Then
                !QuotationAuthorization = 1
        Else: !QuotationAuthorization = 0
        End If

        If frmODASPOperationType.optJobBriefPreparation = True Then
                !JobBriefPreparation = 1
            Else: !JobBriefPreparation = 0
        End If
        
        If frmODASPOperationType.optJobBriefApproval = True Then
                !JobBriefApproval = 1
        Else: !JobBriefApproval = 0
        End If
        
        If frmODASPOperationType.optJobBriefAuthorization = True Then
                !JobBriefAuthorization = 1
        Else: !JobBriefAuthorization = 0
        End If

        If frmODASPOperationType.optJobCardPreparation = True Then
                !JobCardPreparation = 1
            Else: !JobCardPreparation = 0
        End If
        
        If frmODASPOperationType.optJobCardApproval = True Then
                !JobCardApproval = 1
        Else: !JobCardApproval = 0
        End If
        
        If frmODASPOperationType.optJobCardAuthorization = True Then
                !JobCardAuthorization = 1
        Else: !JobCardAuthorization = 0
        End If

        If frmODASPOperationType.optInvoicePreparation = True Then
                !InvoicePreparation = 1
            Else: !InvoicePreparation = 0
        End If
        
        If frmODASPOperationType.optInvoiceApproval = True Then
                !InvoiceApproval = 1
        Else: !InvoiceApproval = 0
        End If
        
        If frmODASPOperationType.optInvoiceAuthorization = True Then
                !InvoiceAuthorization = 1
        Else: !InvoiceAuthorization = 0
        End If

        bSaveRECORD = False
        
         .Update
         .Requery
  End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If

End Sub


Private Sub cmdUpdate_Click()
        bSaveRECORD = True
        ValidateRECORD
        If bSaveRECORD = True Then
            SaveRECORD
                If bSaveRECORD = False Then
'                    enableButtons
                    disableALLRECORD
                    baddRECORD = False
                End If
        End If
        getALLOPERATIONS
        
End Sub

Private Sub cmdSearch_Click()
        searchMyRecord
End Sub

Private Sub Form_Activate()
    disableALLRECORD
'    enableButtons
    
    getALLOPERATIONS
End Sub

Private Sub Form_Load()

    OpenODBCConnection
      
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPOperationType.txtOperationType.Text = Item.Text
            loadRECORD
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub



