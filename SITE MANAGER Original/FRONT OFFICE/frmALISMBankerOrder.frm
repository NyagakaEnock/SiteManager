VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmALISMBankerOrder 
   Caption         =   "Bankers Order Processing"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1111
      ButtonWidth     =   2170
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Loan Details"
            Key             =   "loan"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Policy Details"
            Key             =   "policy"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Premium Details"
            Key             =   "premium"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New"
            Key             =   "add"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "save"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "cancel"
            Key             =   "cancel"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11655
      Begin VB.Frame Frame2 
         Caption         =   "Recent Changes"
         Height          =   1455
         Left            =   120
         TabIndex        =   37
         Top             =   5520
         Width           =   11415
         Begin MSComctlLib.ListView ListView2 
            Height          =   1095
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
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
      Begin VB.Frame Frame4 
         Caption         =   "Policy Details"
         Height          =   1815
         Left            =   4080
         TabIndex        =   23
         Top             =   120
         Width           =   7455
         Begin VB.TextBox txtOrderDate 
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
            Left            =   1440
            TabIndex        =   26
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtOrderAmount 
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
            Left            =   4920
            TabIndex        =   2
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtPolicyNo 
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
            Left            =   4920
            TabIndex        =   25
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtNames 
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
            Left            =   1440
            TabIndex        =   24
            Top             =   840
            Width           =   5895
         End
         Begin VB.TextBox txtOrderNo 
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
            Left            =   1440
            TabIndex        =   1
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Policy Holder"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Order Date"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Order Amount"
            Height          =   255
            Left            =   3720
            TabIndex        =   29
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Policy No"
            Height          =   255
            Left            =   3720
            TabIndex        =   28
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Order No"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Payment Details"
         Height          =   2175
         Left            =   4080
         TabIndex        =   15
         Top             =   1920
         Width           =   7455
         Begin VB.TextBox cboPaymentMode 
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
            Left            =   1440
            TabIndex        =   36
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtPaymentModeDescription 
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
            Left            =   3480
            TabIndex        =   34
            Top             =   1680
            Width           =   3735
         End
         Begin VB.ComboBox cboPaymentMethod 
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
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtPaymentMethodDescription 
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
            Left            =   3480
            TabIndex        =   17
            Top             =   240
            Width           =   3735
         End
         Begin VB.ComboBox cboBankName 
            Appearance      =   0  'Flat
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
            Left            =   2520
            TabIndex        =   4
            Top             =   720
            Width           =   4695
         End
         Begin VB.TextBox txtBankNo 
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
            Left            =   1440
            TabIndex        =   16
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAccountNo 
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
            Left            =   1440
            TabIndex        =   5
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtDateofFirstPayment 
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
            Left            =   5280
            TabIndex        =   6
            Top             =   1200
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPickerDateofFirstPayment 
            Height          =   375
            Left            =   6960
            TabIndex        =   18
            Top             =   1200
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   38015
         End
         Begin VB.Label Label14 
            Caption         =   "Payment Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1740
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Bank No"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Date of First Payment"
            Height          =   255
            Left            =   3600
            TabIndex        =   19
            Top             =   1260
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Company Details"
         Height          =   1455
         Left            =   4080
         TabIndex        =   9
         Top             =   4080
         Width           =   7455
         Begin VB.ComboBox cboCoyBankNO 
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
            Left            =   1440
            TabIndex        =   7
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtCoyBankName 
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
            Left            =   3480
            TabIndex        =   11
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtCoyAccountNo 
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
            Left            =   1440
            TabIndex        =   10
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtIssuedBy 
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
            Left            =   4680
            TabIndex        =   8
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label16 
            Caption         =   "Bank No"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Issued By"
            Height          =   255
            Left            =   3600
            TabIndex        =   12
            Top             =   900
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5295
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9340
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":0542
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":0A84
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":0FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":1508
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":1A4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":1F8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":24CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":2A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":2F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":3494
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":39D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmALISMBankerOrder.frx":3F18
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmALISMBankerOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim translater As New cMoneyConverter
Dim rsbank As clsALISBankerOrder


Private Sub DTPickerDateofFirstPayment_Change()
'On Error GoTo err
    Screen.ActiveForm.DTPickerDateofFirstPayment.MinDate = Date
    Screen.ActiveForm.DTPickerDateofFirstPayment.MousePointer = cc2IBeam
    Screen.ActiveForm.txtDateofFirstPayment.Text = Screen.ActiveForm.DTPickerDateofFirstPayment.Value
Exit Sub
err:
    ErrorMessage
End Sub



Private Sub cboCoybankNo_GotFocus()
'On Error GoTo err

        Dim rsBANKGF As ADODB.Recordset, strBANKGF As String
        Set rsBANKGF = New Recordset
        
        strBANKGF = "SELECT * FROM ALISPBankAccount;"
        rsBANKGF.Open strBANKGF, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        Screen.ActiveForm.cboCoyBankNO.Clear

        With rsBANKGF
            Do Until .EOF
            Screen.ActiveForm.cboCoyBankNO.AddItem !Details
                    .MoveNext
            Loop
    
        End With

rsBANKGF.Close
strBANKGF = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboCoyBankNo_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub cboCoybankNo_LostFocus()
'On Error GoTo err

        Dim rsBANKLF As ADODB.Recordset, strBANKLF As String
        Set rsBANKLF = New Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBankAccount WHERE Details = '" & Screen.ActiveForm.cboCoyBankNO.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboCoyBankNO.Text = !BankNo
                Screen.ActiveForm.txtCoyBankName.Text = !Details
                Screen.ActiveForm.txtCoyAccountNo.Text = !AccountNo
        End With
        
rsBANKLF.Close
strBANKLF = ""

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub Form_Activate()
        Set rsbank = New clsALISBankerOrder
        baddRECORD = True
        rsbank.loadRECORD
        rsbank.getBANKORDER
        rsbank.getPASTBANKORDER
        Set rsbank = Nothing
        
End Sub

Private Sub ListView1_DblClick()
'On Error GoTo err
        With Screen.ActiveForm
                .txtOrderNo.Text = CurrentRecord
                Set rsbank = New clsALISBankerOrder
                baddRECORD = False
                rsbank.loadORDER
                rsbank.loadRECORD
                Set rsbank = Nothing
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        With Screen.ActiveForm
                
                .txtOrderNo.Text = Item.Text
                Set rsbank = New clsALISBankerOrder
                rsbank.loadORDER
                rsbank.loadRECORD
                Set rsbank = Nothing
        End With
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo err
    Set rsbank = New clsALISBankerOrder

    With Screen.ActiveForm
        Select Case (Button.Key)
        
        Case "policy":  Load frmALISMViewPolicy
                        frmALISMViewPolicy.Show 1, Me
        
        Case "loan":    Load Screen.ActiveForm
                        Screen.ActiveForm.Show 1, Me
        
        Case "add":     baddRECORD = True
                        rsbank.addRECORD
        
        Case "save":    rsbank.saveORDER
                        baddRECORD = False
        
        Case "cancel":  rsbank.Cancelrecord
                        baddRECORD = False
        
        Case "print": rsbank.printRECORD
        End Select
    End With
    
    Set rsbank = Nothing

Exit Sub

err:
    ErrorMessage
End Sub




Private Sub loadPaymentModeDESCRIPTION()
'On Error GoTo err

        Dim rsPAYMENTMODELF As ADODB.Recordset
        Set rsPAYMENTMODELF = New Recordset
        
        rsPAYMENTMODELF.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & Screen.ActiveForm.cboPaymentMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPAYMENTMODELF
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.txtPaymentModeDescription.Text = !Description
        End With
  
rsPAYMENTMODELF.Close

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub cboPaymentMethod_GotFocus()
    selectPaymentMethodGotFocus
End Sub

Private Sub cboPaymentMethod_LostFocus()
    selectPaymentMethodLostFocus
End Sub


Private Sub cboBankName_GotFocus()
'On Error GoTo err

        Dim rsBANKGF As ADODB.Recordset, strBANKGF As String
        Set rsBANKGF = New Recordset
        
        strBANKGF = "SELECT * FROM ALISPBank;"
        rsBANKGF.Open strBANKGF, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        Screen.ActiveForm.cboBankName.Clear

        With rsBANKGF
            Do Until .EOF
            Screen.ActiveForm.cboBankName.AddItem !CompanyName
                    .MoveNext
            Loop
    
        End With

rsBANKGF.Close
strBANKGF = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboBankName_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub cboBankName_LostFocus()
'On Error GoTo err

        Dim rsBANKLF As ADODB.Recordset, strBANKLF As String
        Set rsBANKLF = New Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBank WHERE CompanyName= '" & Screen.ActiveForm.cboBankName.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.txtBankNo.Text = !BankNo
                Screen.ActiveForm.cboBankName.Text = !CompanyName
        End With
        
rsBANKLF.Close
strBANKLF = ""

Exit Sub

err:
        ErrorMessage

End Sub




