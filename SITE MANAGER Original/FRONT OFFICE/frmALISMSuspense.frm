VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmALISMSuspense 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A.L.I.S ENTERPRISE [SUSPENSE PROCESSING]"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   1560
   ClientWidth     =   10590
   Icon            =   "frmALISMSuspense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10590
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10455
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   10215
         Begin VB.TextBox txtTermOfPolicy 
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
            Height          =   360
            Left            =   6960
            TabIndex        =   46
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtMaturityDate 
            Alignment       =   2  'Center
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
            Height          =   360
            Left            =   8640
            TabIndex        =   43
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtDateOfCommencement 
            Alignment       =   2  'Center
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
            Height          =   360
            Left            =   5400
            TabIndex        =   41
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtNames 
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
            Height          =   360
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox cboDocumentNo 
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
            Height          =   360
            Left            =   360
            TabIndex        =   37
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Term"
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
            Left            =   7680
            TabIndex        =   45
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Maturity Date"
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
            Left            =   8880
            TabIndex        =   44
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "DOC"
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
            Left            =   5760
            TabIndex        =   42
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Names"
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
            TabIndex        =   40
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblDocumentNo 
            Caption         =   "Document No"
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
            Left            =   720
            TabIndex        =   38
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3135
         Left            =   120
         TabIndex        =   33
         Top             =   3480
         Width           =   8895
         Begin MSComctlLib.ListView ListView1 
            Height          =   2775
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4895
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
      Begin VB.Frame Frame7 
         Height          =   2535
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   7455
         Begin VB.TextBox txtReceiptAMount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   360
            Left            =   6000
            TabIndex        =   60
            Top             =   360
            Width           =   1260
         End
         Begin VB.TextBox txtPaymentStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   1800
            Width           =   1260
         End
         Begin VB.TextBox txtReceivedTodate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   1320
            Width           =   1260
         End
         Begin VB.TextBox txtDateOfLastPayment 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   840
            Width           =   1260
         End
         Begin VB.TextBox txtTransactionAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   360
            Left            =   6000
            TabIndex        =   1
            Top             =   360
            Width           =   1260
         End
         Begin VB.TextBox txtReferenceNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   3240
            TabIndex        =   52
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtAccountingPeriod 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox cboEmployerCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            TabIndex        =   48
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox cboReceiptType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   3240
            TabIndex        =   47
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtExpectedAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtReceiptNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            TabIndex        =   25
            Top             =   345
            Width           =   1335
         End
         Begin VB.TextBox txtReceiptDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   3240
            TabIndex        =   24
            Top             =   345
            Width           =   1455
         End
         Begin VB.TextBox txtEmployeeNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   3240
            TabIndex        =   23
            Top             =   885
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Status"
            Height          =   210
            Left            =   4920
            TabIndex        =   59
            Top             =   1875
            Width           =   615
         End
         Begin VB.Label lblTotalReceived 
            Caption         =   "Total Recvd"
            Height          =   255
            Left            =   4920
            TabIndex        =   57
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label lblDateOfLastPayment 
            Caption         =   "Last Pay Date"
            Height          =   255
            Left            =   4920
            TabIndex        =   55
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Amount"
            Height          =   255
            Left            =   4920
            TabIndex        =   53
            Top             =   420
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Period"
            Height          =   210
            Left            =   120
            TabIndex        =   51
            Top             =   1875
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Ref No"
            Height          =   255
            Left            =   2640
            TabIndex        =   49
            Top             =   1860
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Expected Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Receipt No"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   405
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   " Date"
            Height          =   255
            Left            =   2640
            TabIndex        =   30
            Top             =   405
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "Type"
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   1380
            Width           =   375
         End
         Begin VB.Label Label17 
            Caption         =   "Employer"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Emp #"
            Height          =   255
            Left            =   2640
            TabIndex        =   27
            Top             =   945
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Premium Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   7680
         TabIndex        =   9
         Top             =   960
         Width           =   2655
         Begin VB.TextBox txtUnitCountBeforePayment 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            TabIndex        =   15
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtDueDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtStatusCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            TabIndex        =   13
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtunitsPaid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            TabIndex        =   12
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtAccountNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtSuspenseAccount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1200
            TabIndex        =   10
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblUnitCountBeforePayment 
            Caption         =   "Prem Prior:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1733
            Width           =   855
         End
         Begin VB.Label lblUnitssPaid 
            Caption         =   "Prem Paid:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1373
            Width           =   855
         End
         Begin VB.Label lblStatusCode 
            Caption         =   "Status Code"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1013
            Width           =   975
         End
         Begin VB.Label lblDueDate 
            Caption         =   "Due Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   653
            Width           =   855
         End
         Begin VB.Label lblUnitCount 
            Caption         =   "AcciuntNo"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   293
            Width           =   975
         End
         Begin VB.Label lblSuspenseAccount 
            Caption         =   "Suspense"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2093
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   9120
         TabIndex        =   3
         Top             =   3480
         Width           =   1215
         Begin VB.CommandButton cmdDelete 
            Appearance      =   0  'Flat
            Height          =   400
            Left            =   120
            Picture         =   "frmALISMSuspense.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   2670
            Width           =   975
         End
         Begin VB.CommandButton cmdedit 
            Appearance      =   0  'Flat
            Caption         =   "&Edit"
            Height          =   400
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2265
            Width           =   975
         End
         Begin VB.CommandButton cmdCancel 
            Appearance      =   0  'Flat
            Height          =   400
            Left            =   120
            Picture         =   "frmALISMSuspense.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1455
            Width           =   975
         End
         Begin VB.CommandButton cmdSearch 
            Appearance      =   0  'Flat
            Height          =   400
            Left            =   120
            Picture         =   "frmALISMSuspense.frx":0646
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1050
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            Appearance      =   0  'Flat
            Height          =   400
            Left            =   120
            Picture         =   "frmALISMSuspense.frx":0748
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   645
            Width           =   975
         End
         Begin VB.CommandButton cmdAddNew 
            Appearance      =   0  'Flat
            Height          =   400
            Left            =   120
            Picture         =   "frmALISMSuspense.frx":084A
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdPrint 
            Appearance      =   0  'Flat
            Height          =   400
            Left            =   120
            Picture         =   "frmALISMSuspense.frx":094C
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1860
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmALISMSuspense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSUSPENSE As clsReceipting, rsReceiptDetails As clsSUSPENSE
Dim bunloadFORM As Boolean
Public rsDEP As ADODB.Recordset, strDEP As String

Private Sub Form_Unload(Cancel As Integer)
    If addpen = True Then
        Cancel = True
        MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
    Else
        Cancel = False
    End If
End Sub

Private Sub ClearReceipt()
    Set rsSUSPENSE = New clsReceipting
    rsSUSPENSE.clearRECORD
    Set rsSUSPENSE = Nothing
End Sub

Private Sub disableControls()
        Set rsSUSPENSE = New clsReceipting
        Dim bVAL As Boolean
        rsSUSPENSE.disableDATAENTRY
        Set rsSUSPENSE = Nothing
End Sub

Private Sub UpdatePremium()

    Set rsSUSPENSE = New clsReceipting
    rsSUSPENSE.updateRECORD
    Set rsSUSPENSE = Nothing

End Sub


Private Sub cbobankNo_GotFocus()
    Set rsSUSPENSE = New clsReceipting
    rsSUSPENSE.selectBankNOGotFocus
    Set rsSUSPENSE = Nothing
    
End Sub

Private Sub cboBankNo_KeyPress(KeyAscii As Integer)

        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectBankNoKeyPress (KeyAscii)
        Set rsSUSPENSE = Nothing

End Sub

Private Sub cbobankNo_LostFocus()
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectBankNoLostFocus
        Set rsSUSPENSE = Nothing
End Sub

Private Sub cboCurrencyCode_GotFocus()
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectCURRENCYGOTFOCUS
        Set rsSUSPENSE = Nothing
End Sub

Private Sub cboCurrencyCode_KeyPress(KeyAscii As Integer)
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectCURRENCYKEYPRESS (KeyAscii)
        Set rsSUSPENSE = Nothing
End Sub
Private Sub cboCurrencyCode_LostFocus()
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectCURRENCYLOSTFOCUS
        Set rsSUSPENSE = Nothing
End Sub

Private Sub cboPaymentMethod_GotFocus()
        
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectPaymentMethodGotFocus
        Set rsSUSPENSE = Nothing

End Sub

Private Sub cboPaymentMethod_KeyPress(KeyAscii As Integer)
        
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectPaymentMethodKeyPress (KeyAscii)
        Set rsSUSPENSE = Nothing

End Sub

Private Sub cboPaymentMethod_LostFocus()

        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectPaymentMethodLostFocus
        Set rsSUSPENSE = Nothing

End Sub
Private Sub cboPaymentMethodGotFocus()
        
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectPaymentMethodGotFocus
        Set rsSUSPENSE = Nothing

End Sub

Private Sub cboPaymentMethodKeyPress(KeyAscii As Integer)
        
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectPaymentMethodKeyPress (KeyAscii)
        Set rsSUSPENSE = Nothing
        
End Sub

Private Sub cboPaymentMethodLostFocus()
        
        Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.selectPaymentMethodLostFocus
        Set rsSUSPENSE = Nothing
        
End Sub

Private Sub cmdAddNew_Click()
        enableALLRECORD
        disableButtons
End Sub

Private Sub cmdCancel_Click()
        clearALLRECORD
        disableALLRECORD
        enableButtons
End Sub

Private Sub cmdprintlisting_Click()
    Load frmReceiptListing
    frmReceiptListing.Show 1, Me
End Sub

Private Sub cmdprintreceipt_Click()
        Load frmNewReceipt
        frmNewReceipt.Show 1, Me

End Sub

Private Sub cmdPrint_Click()
    If frmODASMReceipt.txtReceiptNo.Text <= "" Then
        MsgBox "Cannot Use this Form Directly, Load the Receipt on the First Tab", vbOKOnly
        
        Exit Sub
        Else: Load frmNewReceipt
        frmNewReceipt.Show 1, Me
    End If

End Sub

Private Sub cmdSearch_Click()
    Set rsSUSPENSE = New clsReceipting
        rsSUSPENSE.searchRECORD
        If bsearchRECORD = True Then
            rsSUSPENSE.loadRECEIPTDETAILS
            rsSUSPENSE.loadEMPLOYER
            showRECEIPTITEMS
        End If
    Set rsSUSPENSE = Nothing
End Sub

Private Sub cmdUpdate_Click()
    Set rsReceiptDetails = New clsSUSPENSE
        rsReceiptDetails.processUPDATE
    Set rsReceiptDetails = Nothing
    
Exit Sub
End Sub

Private Sub Form_Activate()
        Set rsReceiptDetails = New clsSUSPENSE
        rsReceiptDetails.verifyJOBBRIEFDETAILS
        rsReceiptDetails.LoadDEFAULT
        getRECEIPTS
        Set rsReceiptDetails = Nothing
        enableButtons
        disableALLRECORD
End Sub

Public Sub generateRECEIPTNo()
            Set rsReceiptDetails = New clsReceipting
            rsReceiptDetails.createRECEIPT
            Set rsReceiptDetails = Nothing
End Sub

Private Sub txtTransactionAmount_LostFocus()
            Set rsReceiptDetails = New clsSUSPENSE
            rsReceiptDetails.processReceipt
            Set rsReceiptDetails = Nothing
End Sub

