VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASPLandLord 
   Caption         =   "Land Lord Details"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   240
      TabIndex        =   37
      Top             =   2760
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Select Landlord"
      TabPicture(0)   =   "frmODASMLandLord.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtSearchName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSearch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "View Land Lords Properties"
      TabPicture(1)   =   "frmODASMLandLord.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00808000&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtSearchName 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   10560
         MaxLength       =   5
         TabIndex        =   38
         Top             =   1380
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   120
         TabIndex        =   40
         Top             =   420
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5318
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
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   41
         Top             =   420
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   5318
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
            Name            =   "Arial Narrow"
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
      Caption         =   "Contact Details"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Top             =   6360
      Width           =   12615
      Begin VB.TextBox txtContactName 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtContactDepartment 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   10
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox txtTelephoneExtention 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   7440
         MaxLength       =   15
         TabIndex        =   9
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtContactDesignation 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   11
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label8 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Extension"
         Height          =   255
         Left            =   6240
         TabIndex        =   33
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Designation"
         Height          =   255
         Left            =   6240
         TabIndex        =   32
         Top             =   660
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Addresses"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6240
      TabIndex        =   22
      Top             =   840
      Width           =   6495
      Begin VB.TextBox txtTownDescription 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3120
         TabIndex        =   24
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox cboTownCode 
         BackColor       =   &H00FFC0C0&
         Height          =   345
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtemailAddress 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   6
         Top             =   952
         Width           =   4935
      End
      Begin VB.TextBox txtMobileNo 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtTelephoneNo 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtPostalAddress 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   3
         Top             =   232
         Width           =   1815
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4200
         TabIndex        =   23
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Town Code"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "email Address"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Mobile No"
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Telephone No"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Postal Address"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "Status"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   630
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   6015
      Begin VB.TextBox txtPhysicalAddress 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox txtDateCreated 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtAccountTypeDescription 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2640
         TabIndex        =   16
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtLandLordName 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtLandLordNo 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox cboAccountType 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Physical Address"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Account Type"
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Date Created"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Names"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Land Lord No"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12735
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   660
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   1164
         ButtonWidth     =   3069
         ButtonHeight    =   1005
         TextAlignment   =   1
         ImageList       =   "ImageList1(1)"
         DisabledImageList=   "ImageList1(1)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New &Record"
               Key             =   "N"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Edit/Change "
               Key             =   "E"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Search/Find "
               Key             =   "S"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Refresh "
               Key             =   "R"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print &Preview "
               Key             =   "P"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Help System  "
               Key             =   "H"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
         Begin MSComctlLib.ImageList ImageList1 
            Index           =   1
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   8
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":0038
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":06B2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":0BF4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":1046
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":1360
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":19DA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":2054
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":24A6
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Index           =   0
            Left            =   10920
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   7
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":2B20
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":319A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":35EC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":3906
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":3F80
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":45FA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmODASMLandLord.frx":4A4C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComDlg.CommonDialog HelpCommonDialog 
            Left            =   10560
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            HelpCommand     =   11
            HelpContext     =   72
            HelpFile        =   "REGHELP.HLP"
         End
      End
   End
End
Attribute VB_Name = "frmODASPLandLord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsLANDLORD As clsODASLandLord, MyCommonData As clsCommonData
Private Sub cboTownCode_GotFocus()
With Me
    If .cboTownCode.ListCount <> 0 Then Exit Sub
    AttachSQL = "SELECT Town AS SelectField FROM ODASPTown ORDER BY Town;"
    If .cboTownCode = Empty Then
        .cboTownCode.Clear
    End If
    AttachDropDowns
End With
End Sub

Private Sub cboTownCode_LostFocus()
On Error GoTo err
With Me
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM ODASPTown WHERE Town = '" & .cboTownCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    .txtTownDescription.Text = rsFindRecord!Town
    .cboTownCode.Text = rsFindRecord!TownCode
End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub cmdAddNew_Click()
        NewRecord = True
        enableALLRECORD
        
End Sub
Private Sub cmdDelete_Click()

End Sub

Private Sub saveALTERATION()

On Error GoTo err
        
        Dim rsALTER As ADODB.Recordset, strALTER As String
        Set rsALTER = New Recordset
        
        strALTER = "Select * from ALISMAlterBeneficiary;"
        rsALTER.Open strALTER, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsALTER
            .AddNew
            !PolicyNo = frmALISMBeneficiary.cboPolicyNo
            !Surname = frmALISMBeneficiary.txtSurname
            !othernames = frmALISMBeneficiary.txtOthernames
            !TitleCode = frmALISMBeneficiary.cboTitleCode
            !DateChanged = Date
            !ChangedBy = CurrentUserName
            !BeneficiaryNo = frmALISMBeneficiary.txtBeneficiaryNo & ""
            !RelationshipCode = frmALISMBeneficiary.cboRelationshipCode
            !BeneficiaryType = frmALISMBeneficiary.cboBeneficiaryType
            !Status = frmALISMBeneficiary.cboBeneficiaryStatus
            !LAIdentityNo = frmALISMBeneficiary.txtLAIdentityNo.Text
            .Update
            frmALISMBeneficiary.txtAlterationNo.Text = !AlterationNo
            .Requery
        End With

Exit Sub

err:
    ErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                rsALTER.CancelUpdate
                rsALTER.Requery
        Else
                UpdateErrorMessage
        End If

End Sub
Private Sub validateEdit()
On Error GoTo err
            
            If frmALISMBeneficiary.txtOthernames <= "" Then
                    MsgBox "The Other Names Cannot be Left Blank"
                    frmALISMBeneficiary.txtOthernames.SetFocus
                    
            ElseIf frmALISMBeneficiary.txtBeneficiaryNo.Text <= "" Then
                    MsgBox "The Beneficiary No Cannot be Left Blank"
                    frmALISMBeneficiary.txtBeneficiaryNo.SetFocus
        
            ElseIf frmALISMBeneficiary.txtComment.Text <= "" Then
                    MsgBox "The Comment is required to indicate the Reason for Change of Beneficiary"
                    frmALISMBeneficiary.txtComment.SetFocus
            
            ElseIf frmALISMBeneficiary.cboRelationshipCode.Text <= "" Then
                    MsgBox "The Relationship to the Life Assured is Required"
                    frmALISMBeneficiary.cboRelationshipCode.SetFocus
            
            ElseIf frmALISMBeneficiary.txtLAIdentityNo <= "" Then
                    MsgBox "The Identity of the Life Assured is Required prior to confirming the Change"
                    frmALISMBeneficiary.txtLAIdentityNo.SetFocus
            
            Else
                bEdit = True
                
            End If

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdEdit_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSearch_Click()
        Me.txtSearchName.Locked = False
        
        If Me.cmdSearch.Caption = "Finish" Then
                getLANDLORDS
                bsearchRECORD = False
                Me.cmdSearch.Caption = "Search"
        ElseIf Me.cmdSearch.Caption = "Search" Then
                Me.txtSearchName.Text = Empty
                bsearchRECORD = True
                getLANDLORDS
                Me.cmdSearch.Caption = "Finish"
        End If

End Sub

Private Sub Form_Activate()
        getLANDLORDS
        Set MyCommonData = New clsCommonData
        disableALLRECORD
        rsLANDLORD.LoadDEFAULT
        LoadAccountType

End Sub

Private Sub Form_Initialize()
        Set rsLANDLORD = New clsODASLandLord
End Sub

Private Sub Form_Load()
    getLANDLORDTYPE
End Sub
Private Sub getLANDLORDTYPE()
On Error GoTo err
With Me

        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT *  FROM ODASPAccountType WHERE AccountType = 'LLORD';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        
        .cboAccountType.Text = rsCONTROL!AccountType
        
        Set rsCONTROL = Nothing
        strSQL = Empty

End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub Form_Terminate()
        Set rsLANDLORD = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

        If NewRecord Then
            Cancel = True
            MsgBox "System cannot close this form when registration of Landlord is in Process...", vbCritical + vbOKOnly, "Interruption of Job Brief creation process"
        Else
            Cancel = False
        End If
Exit Sub
err:
ErrorMessage

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
        
        If Item.Checked = True Then
            
                j = Screen.ActiveForm.ListView1.ListItems.Count
                
                If j = 0 Then Exit Sub
                
                For i = 1 To j
                    If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                       Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                    End If
                Next i
                
                frmODASPLandLord.txtLandLordNo = Item.Text
                rsLANDLORD.loadRECORD
                showALLSITESByLandlord
                

        End If

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub txtAccountNo_Change()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "New &Record"
                        NewRecord = True
                        enableALLRECORD
                        Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                        .ListView1.Enabled = False
                    Case "&Save Record"
                        rsLANDLORD.updateRECORD
                        NewRecord = False: Button.Caption = "New &Record": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                        
                        End Select
            Case "E"
                Select Case Button.Caption
                    Case "&Edit/Change "
                         If NewRecord Then Exit Sub
                                If .txtLandLordNo.Text = Empty Then
                                MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                               .txtLandLordNo.SetFocus
                                Else
                                enableALLRECORD
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                EditRecord = True
                            End If
                    Case "Save &Changes "
                        rsLANDLORD.updateRECORD
                        EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 6
                End Select
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                        rsLANDLORD.SearchRecord
                        Button.Caption = "Delete"
                    Case "Cancel"
                        cancelCMD
                    Case "Delete"
                        cmdDelete_Click
                        Button.Caption = "&Search/Find "
                End Select
                .ListView1.Enabled = True
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 3
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 6
                    .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen

            Case "P"
                Load frmODASRLandlordSites
                frmODASRLandlordSites.Show 1, Me

            Case "H"
            .HelpCommonDialog.DialogTitle = "Using the Main System"
            .HelpCommonDialog.HelpFile = App.HelpFile
            .HelpCommonDialog.HelpContext = 18
            .HelpCommonDialog.HelpCommand = cdlHelpContext
            .HelpCommonDialog.ShowHelp

        End Select
    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub txtLandLordName_LostFocus()
On Error GoTo er
    With Me
        .txtContactName.Text = .txtLandLordName.Text
        .txtLandLordName.Text = UCase(.txtLandLordName.Text)
    End With
Exit Sub
er:
ErrorMessage
End Sub
