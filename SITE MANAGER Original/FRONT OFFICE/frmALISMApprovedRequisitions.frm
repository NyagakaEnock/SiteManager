VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmALISMApprovedRequisitions 
   Caption         =   "cheque Processing"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmALISMApprovedRequisitions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Height          =   7215
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   11055
         Begin VB.Frame Frame3 
            Height          =   2535
            Left            =   3840
            TabIndex        =   19
            Top             =   120
            Width           =   7095
            Begin VB.TextBox txtAmountDue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1080
               TabIndex        =   37
               Top             =   1585
               Width           =   2415
            End
            Begin VB.TextBox txtPaymentFlag 
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   6720
               TabIndex        =   35
               Top             =   2040
               Width           =   255
            End
            Begin VB.TextBox txtAmountPaid 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1080
               TabIndex        =   6
               Top             =   2040
               Width           =   2415
            End
            Begin VB.TextBox txtChequeAmount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   4800
               TabIndex        =   32
               Top             =   2040
               Width           =   1335
            End
            Begin VB.TextBox txtRequisitionDate 
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   4800
               TabIndex        =   29
               Top             =   1139
               Width           =   2175
            End
            Begin VB.TextBox txtDocumentNo 
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   4800
               TabIndex        =   28
               Top             =   1596
               Width           =   2175
            End
            Begin VB.ComboBox cboRequisitionNo 
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   4800
               Sorted          =   -1  'True
               TabIndex        =   26
               Top             =   682
               Width           =   2175
            End
            Begin VB.TextBox txtPayeeDetails 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1080
               TabIndex        =   5
               Top             =   1130
               Width           =   2415
            End
            Begin VB.ComboBox cboBankNo 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1080
               Sorted          =   -1  'True
               TabIndex        =   4
               Top             =   675
               Width           =   2415
            End
            Begin VB.TextBox txtChequeDate 
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   4800
               TabIndex        =   21
               Top             =   225
               Width           =   1935
            End
            Begin VB.TextBox txtChequeNo 
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1080
               TabIndex        =   3
               Top             =   225
               Width           =   2415
            End
            Begin MSComCtl2.DTPicker DTPickerChequeDate 
               Height          =   375
               Left            =   6720
               TabIndex        =   20
               Top             =   240
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   393216
               Format          =   19660801
               CurrentDate     =   37945
            End
            Begin VB.Label Label8 
               Caption         =   "Amount Due"
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   1665
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Flag"
               Height          =   255
               Left            =   6240
               TabIndex        =   36
               Top             =   2100
               Width           =   375
            End
            Begin VB.Label Label5 
               Caption         =   "Amount Paid"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   2115
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Chk Amount"
               Height          =   255
               Left            =   3720
               TabIndex        =   33
               Top             =   2115
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "Req Date"
               Height          =   255
               Left            =   3720
               TabIndex        =   31
               Top             =   1155
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Document #"
               Height          =   255
               Left            =   3720
               TabIndex        =   30
               Top             =   1635
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Requisition No"
               Height          =   255
               Left            =   3720
               TabIndex        =   27
               Top             =   728
               Width           =   1095
            End
            Begin VB.Label Label6 
               Caption         =   "Payee "
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Label16 
               Caption         =   "Bank No"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   728
               Width           =   735
            End
            Begin VB.Label Label21 
               Caption         =   "Chk Date"
               Height          =   255
               Left            =   3720
               TabIndex        =   23
               Top             =   278
               Width           =   1215
            End
            Begin VB.Label Label15 
               Caption         =   "Chk No"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   300
               Width           =   1215
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Similar Requisitions"
            Height          =   1455
            Left            =   3840
            TabIndex        =   18
            Top             =   2640
            Width           =   7095
            Begin MSComctlLib.ListView ListView3 
               Height          =   1095
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   1931
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
               BackColor       =   16761024
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame5 
            Height          =   3015
            Left            =   8400
            TabIndex        =   13
            Top             =   4080
            Width           =   2535
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   39
               Top             =   2280
               Width           =   2295
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "&Print Voucher"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   17
               Top             =   1890
               Width           =   2295
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "&Print"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   16
               Top             =   1560
               Width           =   2295
            End
            Begin VB.CommandButton cmdSearch 
               Caption         =   "&Search"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   15
               Top             =   900
               Width           =   2295
            End
            Begin VB.CommandButton cmdAddNew 
               Caption         =   "&Add New"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   0
               Top             =   240
               Width           =   2295
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "&Update"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   8
               Top             =   570
               Width           =   2295
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Cancel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   14
               Top             =   1230
               Width           =   2295
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Included Requisitions"
            Height          =   3015
            Left            =   3840
            TabIndex        =   11
            Top             =   4080
            Width           =   4455
            Begin MSComctlLib.ListView ListView2 
               Height          =   2655
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Approved Requisition"
            Height          =   6975
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   3615
            Begin MSComctlLib.ListView ListView1 
               Height          =   6615
               Left            =   120
               TabIndex        =   2
               Top             =   240
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   11668
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
               BackColor       =   16761024
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
      End
   End
End
Attribute VB_Name = "frmALISMApprovedRequisitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsREQUISITION As clsALISCheque

Private Sub cbobankNo_GotFocus()
        strSQL = "SELECT * FROM ALISPBankAccount"
        bankNoGotFocus
End Sub

Private Sub cboBankNo_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cbobankNo_LostFocus()
        strSQL = "SELECT * FROM ALISPBankAccount WHERE Details = '" & cboBankNo.Text & "'"
        BankNoLostFocus
End Sub

Private Sub cmdAddNew_Click()
        bcheckREQUISITION = False
        bcheckREQUISITION = True
        bcheckREQENTRY = True
        
        If bcheckREQUISITION = True Then
                bcheckREQUISITION = False
                enableALLRECORD
                clearALLRECORD
                disableButtons
                
                Set rsREQUISITION = New clsALISCheque
                rsREQUISITION.initializeRECORD
                Set rsREQUISITION = Nothing
                Exit Sub
        End If
End Sub


Private Sub cmdCancel_Click()
        enableButtons
        clearALLRECORD
        disableButtons
        
End Sub

Private Sub cmdUpdate_Click()
        'validateRECORD
'        If BSave = True Then
'            'GenerateChequeNo
'            updateCHEQUEENTRIES
'            saveRECORD
'            'saveCLAIM
'            LoadGrid
'            loadMGRID
'            BSave = False
'            enableCButtons
'            disableCONTROL
'            frmALISMCheque.SSTabCheque.TabEnabled(1) = True
'            frmALISMCheque.SSTabCheque.TabEnabled(2) = False
'            frmALISMCheque.SSTabCheque.Tab = 1
'            disableENTRY
'            frmALISMCheque.cmdAddNewEntry.SetFocus
'        End If
'
'        Set rsClaimApproval = New clsALISApproval
'        rsClaimApproval.loadAPPROVALDETAILS
'        rsClaimApproval.switchCOMMANDBUTTONS
'        Set rsClaimApproval = Nothing

End Sub
Private Sub cmdEntryUpdate_Click()
        If frmALISMCheque.txtEntryChequeNo.Text <= "" Then
            MsgBox "MUST LOAD The Cheque Before Loading the Entries", vbOKOnly
            Exit Sub
        End If
 
'        validateENTRY
'        If bsaveCHEQUE = True Then
'                SaveChequeENTRIES
'                updateREQUISITION
'                updateTOTAL
'                LoadGrid
'                loadMGRID
'                bsaveCHEQUE = False
'                enableCBEntry
'                disableENTRY
'        End If

End Sub

Private Sub DTPickerChequeDate_Change()
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.ChangeDATE
        Set rsREQUISITION = Nothing

End Sub

Private Sub Form_Activate()
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.GetRequisition
        Set rsREQUISITION = Nothing
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        
        frmALISMCheck.cboRequisitionNo.Text = Item.Text

        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.displayREQUISITION
        Set rsREQUISITION = Nothing
End Sub

Private Sub txtAmountPaid_LostFocus()
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.checkSTATUS
        Set rsREQUISITION = Nothing
End Sub
