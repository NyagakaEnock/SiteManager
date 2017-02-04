VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMReversePayment 
   Caption         =   "Reverse Payments Done"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame15 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtPaymentDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtDueDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAmountPaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtInstallmentNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtAccountNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtContractNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtPaidTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtReferenceNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3120
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPickerVoidDate 
         Height          =   315
         Left            =   8520
         TabIndex        =   15
         Top             =   3540
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55246849
         CurrentDate     =   37953
      End
      Begin VB.Label Label6 
         Caption         =   "Paid On"
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Payment Due Date"
         Height          =   255
         Left            =   5040
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Contract No"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCurrentPeriod 
         Caption         =   "Installment"
         Height          =   210
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Paid To A/C No"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Rent Paid"
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblPaymentStatus 
         Caption         =   "Status"
         Height          =   210
         Left            =   5760
         TabIndex        =   16
         Top             =   2715
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contract Installments"
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6855
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3413
         View            =   3
         MultiSelect     =   -1  'True
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   6960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
      Begin VB.CommandButton cmdNew 
         Height          =   375
         Left            =   120
         Picture         =   "frmODASMReversePayment.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         Picture         =   "frmODASMReversePayment.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         Picture         =   "frmODASMReversePayment.frx":076C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         Picture         =   "frmODASMReversePayment.frx":086E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmODASMReversePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsREVERSE As clsALISReceipt
Public rsreceipt As clsReceipting

Private Sub cmdAddNew_Click()
        Set rsREVERSE = New clsALISReceipt
        rsREVERSE.addRECORD
        Set reverse = Nothing
End Sub
Private Sub cboReversalType_GotFocus()
'        selectReversalGotFocus
End Sub

Private Sub cboReversalType_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboReversalType_LostFocus()
'        selectReversalLostFocus
End Sub

Private Sub cmdCancel_Click()
        Set rsREVERSE = New clsALISReceipt
        rsREVERSE.Cancelrecord
        Set reverse = Nothing
End Sub

Private Sub locateRECORD()
''''On Error GoTo Myerr
With Me

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        rsFind.Open "SELECT * FROM ALISMReceiptNew WHERE ReceiptNo = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
            If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
            Else:
                    Me.txtReferenceNo.Text = !ReceiptNo
                    
            End If
        End With
End With
Exit Sub

Myerr:
    ErrorMessage
End Sub

Private Sub cmdNew_Click()
With Me
    .cmdUpdate.Enabled = True
        
End With
End Sub

Private Sub cmdUpdate_Click()
      ValidateData
        Set rsREVERSE = New clsALISReceipt
        saveINSTALLMENTISSUED
       ' ReverseVoucher
        Set reverse = Nothing
End Sub

Private Sub DTPickerVoidDate_Change()
On Error GoTo err
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableSButtons
    GetVoucherToReverse
End Sub
Public Sub saveINSTALLMENTISSUED()
On Error GoTo err

    With frmODASMReversePayment

                Set rsSAVE = New Recordset
                strSQL = "SELECT * from ODASMInstallment where Invoiceno = '" & .txtInstallmentNo & " ' and  ContractNo = '" & .txtContractNo.Text & " '"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                Else
                        rsSAVE!Status = "REVERSED"
                        rsSAVE!StatusDate = Date
                        rsSAVE!PaymentFlag = "N"
                        rsSAVE!PaymentDue = rsSAVE!TotalRent
                        rsSAVE!Balance = rsSAVE!TotalRent
                        rsSAVE!AmountPaid = CDbl(rsSAVE!TotalRent)
                        rsSAVE!Requisitioned = "N"
                                               
                        rsSAVE.Update
                End If
    End With
    
    Exit Sub
            
        rsCONTROL.Close
        strSQL = ""
    
   
err:
    UpdateErrorMessage
End Sub
Public Sub ReverseVoucher()
On Error GoTo err

    With frmODASMReversePayment

                Set rsSAVE = New Recordset
                strSQL = "SELECT * from ODASMVoucher where ContractNo = '" & .txtContractNo.Text & " '"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                Else
                        rsSAVE!Status = "REVERSED"
                        rsSAVE!StatusDate = Date
                        rsSAVE!PaymentFlag = "N"
                        rsSAVE!PaymentDue = rsSAVE!TotalRent
                        rsSAVE!Balance = rsSAVE!TotalRent
                        rsSAVE!AmountPaid = CDbl(rsSAVE!TotalRent)
                                               
                        rsSAVE.Update
                End If
    End With
    
    Exit Sub
            
        rsCONTROL.Close
        strSQL = ""
    
   
err:
    UpdateErrorMessage
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
With Me
On Error GoTo err
        
    Dim i, j As Double
    
    If Item.Checked = True Then
        
        j = Screen.ActiveForm.ListView1.ListItems.Count
        
        If j = 0 Then Exit Sub
        
        For i = 1 To j
            If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                Screen.ActiveForm.ListView1.ListItems(i).Checked = False
            End If
        Next i
        
        With frmODASMReversePayment
            If .ListView1.Checkboxes = False Then
                    Exit Sub
            End If
        
           .txtInstallmentNo.Text = Item()
           .txtAccountNo.Text = Item.SubItems(4)
           .txtAmountPaid.Text = Item.SubItems(5)
           .txtContractNo.Text = Item.SubItems(1)
           .txtDueDate.Text = Item.SubItems(8)
           .txtPaidTo.Text = Item.SubItems(5)
           .txtPaymentDate = Item.SubItems(8)
  End With
    
    
    End If
  
Exit Sub

err:
    ErrorMessage
End With
End Sub
Public Sub enableSButtons()
On Error GoTo err

    With Me
            .cmdUpdate.Enabled = False
            .cmdNew.Enabled = True
            .cmdCancel.Enabled = True
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ValidateData()

On Error GoTo err

    bsaveRECORD = False
    
    With frmODASMReversePayment
    
              If .txtInstallmentNo.Text = "" Then
                      MsgBox "You MUST check the entry you want to reverse from the list"
                      .txtInstallmentNo.SetFocus
                                   
              ElseIf .txtContractNo.Text = "" Then
                      MsgBox "Contract No. cant be blank"
                      .txtContractNo.SetFocus
            
              ElseIf .txtAccountNo.Text = "" Then
                      MsgBox "Account No. cant be left Blank"
                      .txtAccountNo.SetFocus
                         
              Else
                         bsaveRECORD = True
              End If
    End With
        
                    
Exit Sub

err:
    UpdateErrorMessage
            
End Sub
