VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASmInvoiceIssuance 
   Caption         =   "View Invoice Details"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.Frame Frame8 
         Caption         =   "Invoice Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   8175
         Begin VB.TextBox txtInvoiceDescription 
            BackColor       =   &H00FFFFC0&
            Height          =   1035
            Left            =   1560
            MaxLength       =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   840
            Width           =   6375
         End
         Begin VB.TextBox txtInvoiceDate 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            TabIndex        =   19
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtInvoiceNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1560
            TabIndex        =   18
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Description"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1230
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Invoice Date"
            Height          =   255
            Left            =   4800
            TabIndex        =   22
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Invoice No"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   390
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Client Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   8175
         Begin VB.TextBox txtProductCode 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6480
            TabIndex        =   29
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtDescriptionOfOrder 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   2880
            TabIndex        =   26
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox txtJobBriefNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   840
            TabIndex        =   24
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCurrentPeriod 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6480
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtAccountNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   840
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtCompanyName 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   2040
            TabIndex        =   13
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label6 
            Caption         =   "Period"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   5640
            TabIndex        =   30
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Desc"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   2160
            TabIndex        =   28
            Top             =   750
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Product"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   5640
            TabIndex        =   27
            Top             =   750
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Brief No"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   750
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   " Name"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Invoice Sent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   8175
         Begin MSComctlLib.ListView ListView3 
            Height          =   2055
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3625
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   5640
         Width           =   8175
         Begin VB.TextBox txtprinted 
            Height          =   285
            Left            =   360
            TabIndex        =   32
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00C0C000&
            Caption         =   "PRINT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtRemark 
            BackColor       =   &H00FFFFC0&
            Height          =   555
            Left            =   840
            MaxLength       =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtPriceInclusive 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            TabIndex        =   4
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtVATAmount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            TabIndex        =   3
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtPriceExclusive 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label30 
            Caption         =   "Remark"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Amount Incl"
            Height          =   255
            Left            =   4680
            TabIndex        =   8
            Top             =   1230
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "VAT Amount"
            Height          =   255
            Left            =   4680
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "InVoice Amount"
            Height          =   255
            Left            =   4680
            TabIndex        =   6
            Top             =   270
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmODASmInvoiceIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdPrint_Click()
On Error GoTo err
    With frmODASmInvoiceIssuance
        
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMInvoiceSent WHERE invoiceNo = '" & .txtInvoiceNo.Text & "' ; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        rsCONTROL!Printed = "Y"
        rsCONTROL!DatePrinted = Date
        rsCONTROL.Update
        rsCONTROL.Requery
              
        CurrentRecord = .txtInvoiceNo
        Load frmODASRInvoice
        frmODASRInvoice.Show 1, Me
        
    End With
        
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        disableALLRECORD
        loadRECORD
        showINVOICEDETAILS
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Public Sub loadRECORD()
On Error GoTo err
    With frmODASmInvoiceIssuance
        
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMJobBrief JB, ODASPAccount AC,  ODASMInvoiceSent INV WHERE inv.AccountNo = Ac.AccountNo and JB.JobBriefNo = inv.JobBriefNo and inv.invoiceNo = '" & .txtInvoiceNo.Text & "' ; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                
                .txtAccountNo.Text = rsCONTROL!AccountNo
                .txtCompanyName.Text = rsCONTROL!CompanyName
                .txtDescriptionOfOrder.Text = rsCONTROL!descriptionOfOrder
                .txtProductCode.Text = rsCONTROL!ProductCode
                .txtInvoiceDate.Text = rsCONTROL!InvoiceDate
                .txtInvoiceDescription.Text = rsCONTROL!InvoiceDescription
                .txtJobBriefNo.Text = rsCONTROL!JobBriefNo
                .txtCurrentPeriod.Text = rsCONTROL!CurrentPeriod
                .txtPriceExclusive.Text = rsCONTROL!PriceExclusive
                .txtPriceInclusive.Text = rsCONTROL!PriceInclusive
                .txtVATAmount.Text = rsCONTROL!VATAmount
                .txtRemark.Text = rsCONTROL!remark
                .txtprinted.Text = rsCONTROL!Printed & ""
    End With

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Unload(cancel As Integer)
        If NewRecord = True Then
            cancel = True
            MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
        Else
            cancel = False
        End If

End Sub
