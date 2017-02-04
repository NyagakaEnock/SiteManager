VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmALISPLoanOperationType 
   Caption         =   "Loan Operation Type"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11760
   Icon            =   "frmALISPLoanOperationType.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmALISPLoanOperationType.frx":0442
   ScaleHeight     =   7485
   ScaleWidth      =   11760
   Begin TabDlg.SSTab SSTabSecurity 
      Height          =   7215
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12726
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Setup"
      TabPicture(0)   =   "frmALISPLoanOperationType.frx":0784
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Grid"
      TabPicture(1)   =   "frmALISPLoanOperationType.frx":07A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame14 
         Height          =   6735
         Left            =   -74880
         TabIndex        =   61
         Top             =   360
         Width           =   11295
         Begin MSDataGridLib.DataGrid OperationTypeGrid 
            Height          =   6375
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   11245
            _Version        =   393216
            BackColor       =   16777152
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame12 
         Height          =   6735
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   11415
         Begin VB.Frame frabrowse 
            Height          =   735
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   5760
            Width           =   9015
            Begin VB.CommandButton cmdFirstCode 
               BackColor       =   &H00E0E0E0&
               Height          =   375
               Index           =   0
               Left            =   1200
               Picture         =   "frmALISPLoanOperationType.frx":07BC
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton cmdFirstCode 
               BackColor       =   &H00E0E0E0&
               Height          =   375
               Index           =   1
               Left            =   2640
               Picture         =   "frmALISPLoanOperationType.frx":0BFE
               Style           =   1  'Graphical
               TabIndex        =   59
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton cmdFirstCode 
               BackColor       =   &H00E0E0E0&
               Height          =   375
               Index           =   2
               Left            =   4080
               Picture         =   "frmALISPLoanOperationType.frx":1040
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton cmdFirstCode 
               BackColor       =   &H00E0E0E0&
               Height          =   375
               Index           =   3
               Left            =   5400
               Picture         =   "frmALISPLoanOperationType.frx":1482
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame Frame13 
            Height          =   5655
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   8775
            Begin VB.TextBox txtOperationType 
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
               Height          =   375
               Left            =   2040
               TabIndex        =   1
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtDescription 
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
               Height          =   375
               Left            =   3360
               TabIndex        =   2
               Top             =   240
               Width           =   5295
            End
            Begin VB.Frame Frame4 
               Caption         =   "Cheque"
               Height          =   4815
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   8535
               Begin VB.OptionButton optchequeApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   44
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.OptionButton optChequeAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   43
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.OptionButton optChequePreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   42
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.OptionButton optChequeIssuance 
                  Caption         =   "Issuance"
                  Height          =   255
                  Left            =   6120
                  TabIndex        =   41
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.OptionButton optPaymentApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.OptionButton optPaymentAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   39
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.OptionButton optPaymentPreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   38
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.OptionButton optDischargeApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   37
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.OptionButton optDischargeAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   36
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.OptionButton optDischargePreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   35
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.OptionButton optChecked 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   34
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.OptionButton optAuthorized 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   33
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.OptionButton optApproved 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   32
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.OptionButton optClaimRegPreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   31
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.OptionButton optClaimRegAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   30
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.OptionButton optClaimRegApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   29
                  Top             =   1800
                  Width           =   1095
               End
               Begin VB.OptionButton optReinstatementApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   28
                  Top             =   2160
                  Width           =   1095
               End
               Begin VB.OptionButton optReinstatementAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   27
                  Top             =   2160
                  Width           =   1335
               End
               Begin VB.OptionButton optReinstatementPreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   26
                  Top             =   2160
                  Width           =   1335
               End
               Begin VB.OptionButton optPaidupPreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   25
                  Top             =   2520
                  Width           =   1335
               End
               Begin VB.OptionButton optPaidupAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   24
                  Top             =   2520
                  Width           =   1335
               End
               Begin VB.OptionButton optPaidupApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   23
                  Top             =   2520
                  Width           =   1095
               End
               Begin VB.OptionButton optProposalPreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   22
                  Top             =   2880
                  Width           =   1335
               End
               Begin VB.OptionButton optProposalAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   21
                  Top             =   2880
                  Width           =   1335
               End
               Begin VB.OptionButton optProposalApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   20
                  Top             =   2880
                  Width           =   1095
               End
               Begin VB.OptionButton optPolicyPreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   19
                  Top             =   3240
                  Width           =   1335
               End
               Begin VB.OptionButton optPolicyAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   18
                  Top             =   3240
                  Width           =   1335
               End
               Begin VB.OptionButton optPolicyApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   17
                  Top             =   3240
                  Width           =   1095
               End
               Begin VB.OptionButton optMedicalPreparation 
                  Caption         =   "Preparation"
                  Height          =   255
                  Left            =   2280
                  TabIndex        =   16
                  Top             =   3600
                  Width           =   1335
               End
               Begin VB.OptionButton optMedicalAuthorization 
                  Caption         =   "Authorization"
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   15
                  Top             =   3600
                  Width           =   1335
               End
               Begin VB.OptionButton optMedicalApproval 
                  Caption         =   "Approval"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   14
                  Top             =   3600
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  Caption         =   "Cheque"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   54
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  Caption         =   "Payment"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   53
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.Label Label3 
                  Caption         =   "Loan"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   52
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  Caption         =   "Discharge"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   51
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.Label Label5 
                  Caption         =   "Claim Registration"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   50
                  Top             =   1800
                  Width           =   2055
               End
               Begin VB.Label Label6 
                  Caption         =   "Reinstatement"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   49
                  Top             =   2160
                  Width           =   2055
               End
               Begin VB.Label Label7 
                  Caption         =   "Paidup Processing"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   48
                  Top             =   2520
                  Width           =   2055
               End
               Begin VB.Label Label8 
                  Caption         =   "Create Proposal"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   47
                  Top             =   2880
                  Width           =   2055
               End
               Begin VB.Label Label9 
                  Caption         =   "Create Policy"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   46
                  Top             =   3240
                  Width           =   2055
               End
               Begin VB.Label Label10 
                  Caption         =   "Medical"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   45
                  Top             =   3600
                  Width           =   2055
               End
            End
            Begin VB.Label lblRelationshipCode 
               Caption         =   "OperationType "
               Height          =   255
               Left            =   240
               TabIndex        =   55
               Top             =   315
               Width           =   1335
            End
         End
         Begin VB.Frame fraCButtons 
            Height          =   3855
            Index           =   6
            Left            =   9000
            TabIndex        =   5
            Top             =   120
            Width           =   2295
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
               Height          =   495
               Left            =   120
               TabIndex        =   11
               Top             =   1230
               Width           =   2055
            End
            Begin VB.CommandButton cmdAdd 
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
               Height          =   495
               Left            =   120
               TabIndex        =   0
               Top             =   240
               Width           =   2055
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "&Edit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   10
               Top             =   1725
               Width           =   2055
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
               Height          =   495
               Left            =   120
               TabIndex        =   9
               Top             =   735
               Width           =   2055
            End
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
               Height          =   495
               Left            =   120
               TabIndex        =   8
               Top             =   2220
               Width           =   2055
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
               Height          =   495
               Left            =   120
               TabIndex        =   7
               Top             =   2715
               Width           =   2055
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
               Height          =   495
               Left            =   120
               TabIndex        =   6
               Top             =   3240
               Width           =   2055
            End
         End
      End
   End
End
Attribute VB_Name = "frmALISPLoanOperationType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSAVE As ADODB.Recordset, strRELN As String
Sub LoadGrid()

        Set OperationTypeGrid.DataSource = rsSAVE
End Sub

Sub clearRELN()
        With frmALISPLoanOperationType
            .txtOperationType.Text = ""
            .txtDescription.Text = ""
            .optChecked.Value = 0
            .optApproved.Value = 0
            .optAuthorized.Value = 0
            .optDischargeApproval.Value = 0
            .optDischargeAuthorization.Value = 0
            .optDischargePreparation.Value = 0
            .optPaymentApproval.Value = 0
            .optPaymentAuthorization.Value = 0
            .optPaymentPreparation.Value = 0
            .optchequeApproval.Value = 0
            .optChequeAuthorization.Value = 0
            .optChequeIssuance.Value = 0
            .optChequePreparation.Value = 0
            .optClaimRegApproval.Value = 0
            .optClaimRegAuthorization.Value = 0
            .optClaimRegPreparation.Value = 0
            .optReinstatementApproval.Value = 0
            .optReinstatementAuthorization.Value = 0
            .optReinstatementPreparation.Value = 0
        End With
End Sub

Sub enableRELN()
        With frmALISPLoanOperationType
            .txtOperationType.Locked = False
            .txtDescription.Locked = False
            .optChecked.Enabled = True
            .optApproved.Enabled = True
            .optAuthorized.Enabled = True
            .optDischargeApproval.Enabled = True
            .optDischargeAuthorization.Enabled = True
            .optDischargePreparation.Enabled = True
            .optPaymentApproval.Enabled = True
            .optPaymentAuthorization.Enabled = True
            .optPaymentPreparation.Enabled = True
            .optchequeApproval.Enabled = True
            .optChequeAuthorization.Enabled = True
            .optPaymentPreparation.Enabled = True
            .optchequeApproval.Enabled = True
            .optChequeAuthorization.Enabled = True
            .optChequeIssuance.Enabled = True
            .optChequePreparation.Enabled = True
            .optClaimRegApproval.Enabled = True
            .optClaimRegAuthorization.Enabled = True
            .optClaimRegPreparation.Enabled = True
            .optReinstatementApproval.Enabled = True
            .optReinstatementAuthorization.Enabled = True
            .optReinstatementPreparation.Enabled = True


        End With
End Sub

Sub disableRELN()
        With frmALISPLoanOperationType
                .txtOperationType.Locked = True
                .txtDescription.Locked = True
                .optChecked.Enabled = False
                .optApproved.Enabled = False
                .optAuthorized.Enabled = False
                .optDischargeApproval.Enabled = False
                .optDischargeAuthorization.Enabled = False
                .optDischargePreparation.Enabled = False
                .optPaymentApproval.Enabled = False
                .optPaymentAuthorization.Enabled = False
                .optPaymentPreparation.Enabled = False
                .optchequeApproval.Enabled = False
                .optChequeAuthorization.Enabled = False
                .optChequeIssuance.Enabled = False
                .optChequePreparation.Enabled = False
                .optClaimRegApproval.Enabled = False
                .optClaimRegAuthorization.Enabled = False
                .optClaimRegPreparation.Enabled = False
                
                .optReinstatementApproval.Enabled = False
                .optReinstatementAuthorization.Enabled = False
                .optReinstatementPreparation.Enabled = False

        End With
End Sub

Sub showRELN()
    With rsSAVE
        frmALISPLoanOperationType.txtOperationType = !OperationType
        frmALISPLoanOperationType.txtDescription = !Description
                            
        If !MedicalApproval = True Then
                frmALISPLoanOperationType.optMedicalApproval.Value = True
        Else: frmALISPLoanOperationType.optMedicalApproval.Value = False
        End If
        
        If !MedicalPreparation = True Then
                frmALISPLoanOperationType.optMedicalPreparation.Value = True
        Else: frmALISPLoanOperationType.optMedicalPreparation.Value = False
        End If
        
        If !MedicalAuthorization = True Then
                frmALISPLoanOperationType.optMedicalAuthorization.Value = True
        Else: frmALISPLoanOperationType.optMedicalAuthorization.Value = False
        End If
    
        
        If !ProposalApproval = True Then
                frmALISPLoanOperationType.optProposalApproval.Value = True
        Else: frmALISPLoanOperationType.optProposalApproval.Value = False
        End If
        
        If !ProposalPreparation = True Then
                frmALISPLoanOperationType.optProposalPreparation.Value = True
        Else: frmALISPLoanOperationType.optProposalPreparation.Value = False
        End If
        
        If !ProposalAuthorization = True Then
                frmALISPLoanOperationType.optProposalAuthorization.Value = True
        Else: frmALISPLoanOperationType.optProposalAuthorization.Value = False
        End If
    
        
        
        If !PolicyApproval = True Then
                frmALISPLoanOperationType.optPolicyApproval.Value = True
        Else: frmALISPLoanOperationType.optPolicyApproval.Value = False
        End If
        
        If !PolicyPreparation = True Then
                frmALISPLoanOperationType.optPolicyPreparation.Value = True
        Else: frmALISPLoanOperationType.optPolicyPreparation.Value = False
        End If
        
        If !PolicyAuthorization = True Then
                frmALISPLoanOperationType.optPolicyAuthorization.Value = True
        Else: frmALISPLoanOperationType.optPolicyAuthorization.Value = False
        End If
    
        
        
        
        If !paidupApproval = True Then
                frmALISPLoanOperationType.optPaidupApproval.Value = True
        Else: frmALISPLoanOperationType.optPaidupApproval.Value = False
        End If
        
        If !paidupPreparation = True Then
                frmALISPLoanOperationType.optPaidupPreparation.Value = True
        Else: frmALISPLoanOperationType.optPaidupPreparation.Value = False
        End If
        
        If !paidupAuthorization = True Then
                frmALISPLoanOperationType.optPaidupAuthorization.Value = True
        Else: frmALISPLoanOperationType.optPaidupAuthorization.Value = False
        End If
        
        If !ReinstatementApproval = True Then
                frmALISPLoanOperationType.optReinstatementApproval.Value = True
        Else: frmALISPLoanOperationType.optReinstatementApproval.Value = False
        End If
        
        If !ReinstatementPreparation = True Then
                frmALISPLoanOperationType.optReinstatementPreparation.Value = True
        Else: frmALISPLoanOperationType.optReinstatementPreparation.Value = False
        End If
        
        If !ReinstatementAuthorization = True Then
                frmALISPLoanOperationType.optReinstatementAuthorization.Value = True
        Else: frmALISPLoanOperationType.optReinstatementAuthorization.Value = False
        End If

        
        
        If !RegistrationPreparation = True Then
              frmALISPLoanOperationType.optClaimRegPreparation.Value = 1
        Else: frmALISPLoanOperationType.optClaimRegPreparation.Value = 0
        End If
            
        If !RegistrationApproval = True Then
              frmALISPLoanOperationType.optClaimRegApproval.Value = 1
        Else: frmALISPLoanOperationType.optClaimRegApproval.Value = 0
        End If
        
        If !RegistrationAuthorization = True Then
              frmALISPLoanOperationType.optClaimRegAuthorization.Value = 1
        Else: frmALISPLoanOperationType.optChequeAuthorization.Value = 0
        End If
       
        If !Checked = True Then
              frmALISPLoanOperationType.optChecked.Value = 1
        Else: frmALISPLoanOperationType.optChecked.Value = 0
        End If
            
        If !Approved = True Then
              frmALISPLoanOperationType.optApproved.Value = 1
        Else: frmALISPLoanOperationType.optApproved.Value = 0
        End If
        
        If !authorized = True Then
              frmALISPLoanOperationType.optAuthorized.Value = 1
        Else: frmALISPLoanOperationType.optAuthorized.Value = 0
        End If
            
        If !DischargeApproval = True Then
                frmALISPLoanOperationType.optDischargeApproval.Value = 1
        Else: frmALISPLoanOperationType.optDischargeApproval.Value = 0
        End If
        
        If !DischargePreparation = True Then
                frmALISPLoanOperationType.optDischargePreparation.Value = 1
        Else: frmALISPLoanOperationType.optDischargePreparation.Value = 0
        End If
        
        If !DischargeAuthorization = True Then
                frmALISPLoanOperationType.optDischargeAuthorization.Value = 1
        Else: frmALISPLoanOperationType.optDischargeAuthorization.Value = 0
        End If

        If !PaymentApproval = True Then
                frmALISPLoanOperationType.optPaymentApproval.Value = 1
        Else: frmALISPLoanOperationType.optPaymentApproval.Value = 0
        End If
        
        If !PaymentPreparation = True Then
                frmALISPLoanOperationType.optPaymentPreparation.Value = 1
        Else: frmALISPLoanOperationType.optPaymentPreparation.Value = 0
        End If
        
        If !PaymentAuthorization = True Then
                frmALISPLoanOperationType.optPaymentAuthorization.Value = 1
        Else: frmALISPLoanOperationType.optPaymentAuthorization.Value = 0
        End If
        
        If !ChequeApproval = True Then
                frmALISPLoanOperationType.optchequeApproval.Value = 1
        Else: frmALISPLoanOperationType.optchequeApproval.Value = 0
        End If
        
        If !ChequePreparation = True Then
                frmALISPLoanOperationType.optChequePreparation.Value = 1
        Else: frmALISPLoanOperationType.optChequePreparation.Value = 0
        End If
        
        If !ChequeAuthorization = True Then
                frmALISPLoanOperationType.optChequeAuthorization.Value = 1
        Else: frmALISPLoanOperationType.optChequeAuthorization.Value = 0
        End If
        
        If !ChequeIssuance = True Then
                frmALISPLoanOperationType.optChequeIssuance.Value = 1
        Else: frmALISPLoanOperationType.optChequeIssuance.Value = 0
        End If


    End With
End Sub

Private Sub DisableCButtons()
        With frmALISPLoanOperationType
            .cmdUpdate.Enabled = True
            .cmdadd.Enabled = False
            .cmdSearch.Enabled = False
            .cmdEdit.Enabled = False
            .cmdDelete.Enabled = False
            .cmdCancel.Enabled = True
        End With
End Sub

Private Sub enableCButtons()
    With frmALISPLoanOperationType
            .cmdUpdate.Enabled = False
            .cmdadd.Enabled = True
            .cmdSearch.Enabled = True
            .cmdEdit.Enabled = True
            .cmdDelete.Enabled = True
            .cmdCancel.Enabled = True
    End With
End Sub

Private Sub cmdAdd_Click()
        clearALLRECORD
        enableALLRECORD
        disableButtons
End Sub

Private Sub cmdCancel_Click()
        enableButtons
        clearRELN
        disableRELN
End Sub


Private Sub cmdDelete_Click()
On Error GoTo err

If txtOperationType.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            With rsSAVE
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearRELN
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
On Error GoTo err

Dim strQRE As Variant
Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset

        Select Case cmdEdit.Caption
                Case "&Edit"
                        enableRELN

                        strQRE = InputBox("Enter The Operation Type to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ALISPLoanOperationType WHERE OperationType LIKE '" & strQRE & "';", cnALIS, adOpenKeyset, adLockOptimistic

                        With rsFind
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                        frmALISPLoanOperationType.txtOperationType = !OperationType
                                        frmALISPLoanOperationType.optChecked = !Checked
                                        frmALISPLoanOperationType.optApproved = !Approved
                                        frmALISPLoanOperationType.txtDescription = !Description
                                        frmALISPLoanOperationType.optAuthorized = !authorized
                                                If frmALISPLoanOperationType.optMedicalPreparation = True Then
                                                !MedicalPreparation = 1
                                            Else: !MedicalPreparation = 0
                                        End If
                                        
                                        If frmALISPLoanOperationType.optMedicalApproval = True Then
                                                !MedicalApproval = 1
                                        Else: !MedicalApproval = 0
                                        End If
                                        
                                        If frmALISPLoanOperationType.optMedicalAuthorization = True Then
                                                !MedicalAuthorization = 1
                                        Else: !MedicalAuthorization = 0
                                        End If
                                
                                        If frmALISPLoanOperationType.optProposalPreparation = True Then
                                                !ProposalPreparation = 1
                                            Else: !ProposalPreparation = 0
                                        End If
                                        
                                        If frmALISPLoanOperationType.optProposalApproval = True Then
                                                !ProposalApproval = 1
                                        Else: !ProposalApproval = 0
                                        End If
                                        
                                        If frmALISPLoanOperationType.optProposalAuthorization = True Then
                                                !ProposalAuthorization = 1
                                        Else: !ProposalAuthorization = 0
                                        End If
                                
                                        
                                        If frmALISPLoanOperationType.optPolicyPreparation = True Then
                                                !PolicyPreparation = 1
                                            Else: !PolicyPreparation = 0
                                        End If
                                        
                                        If frmALISPLoanOperationType.optPolicyApproval = True Then
                                                !PolicyApproval = 1
                                        Else: !PolicyApproval = 0
                                        End If
                                        
                                        If frmALISPLoanOperationType.optPolicyAuthorization = True Then
                                                !PolicyAuthorization = 1
                                        Else: !PolicyAuthorization = 0
                                        End If
                                
                                        
                                        
                                        If frmALISPLoanOperationType.optPaidupPreparation = True Then
                                                !paidupPreparation = 1
                                            Else: !paidupPreparation = 0
                                        End If
                                        
                                        If frmALISPLoanOperationType.optPaidupApproval = True Then
                                                !paidupApproval = 1
                                        Else: !paidupApproval = 0
                                        End If
                                        
                                        If frmALISPLoanOperationType.optPaidupAuthorization = True Then
                                                !paidupAuthorization = 1
                                        Else: !paidupAuthorization = 0
                                        End If

                                        If !ReinstatementApproval = True Then
                                                frmALISPLoanOperationType.optReinstatementApproval.Value = True
                                        Else: frmALISPLoanOperationType.optReinstatementApproval.Value = False
                                        End If
                                        
                                        If !ReinstatementPreparation = True Then
                                                frmALISPLoanOperationType.optReinstatementPreparation.Value = True
                                        Else: frmALISPLoanOperationType.optReinstatementPreparation.Value = False
                                        End If
                                        
                                        If !ReinstatementAuthorization = True Then
                                                frmALISPLoanOperationType.optReinstatementAuthorization.Value = True
                                        Else: frmALISPLoanOperationType.optReinstatementAuthorization.Value = False
                                        End If

                                        If !RegistrationApproval = True Then
                                                frmALISPLoanOperationType.optClaimRegApproval.Value = 1
                                        Else: frmALISPLoanOperationType.optClaimRegApproval.Value = 0
                                        End If
                                        
                                        If !RegistrationPreparation = True Then
                                                frmALISPLoanOperationType.optClaimRegPreparation.Value = 1
                                        Else: frmALISPLoanOperationType.optClaimRegPreparation.Value = 0
                                        End If
                                        
                                        If !RegistrationAuthorization = True Then
                                                frmALISPLoanOperationType.optClaimRegAuthorization.Value = 1
                                        Else: frmALISPLoanOperationType.optClaimRegAuthorization.Value = 0
                                        End If

                                        If !DischargeApproval = True Then
                                                frmALISPLoanOperationType.optDischargeApproval.Value = 1
                                        Else: frmALISPLoanOperationType.optDischargeApproval.Value = 0
                                        End If
                                        
                                        If !DischargePreparation = True Then
                                                frmALISPLoanOperationType.optDischargePreparation.Value = 1
                                        Else: frmALISPLoanOperationType.optDischargePreparation.Value = 0
                                        End If
                                        
                                        If !DischargeAuthorization = True Then
                                                frmALISPLoanOperationType.optDischargeAuthorization.Value = 1
                                        Else: frmALISPLoanOperationType.optDischargeAuthorization.Value = 0
                                        End If
                                
                                        If !PaymentApproval = True Then
                                                frmALISPLoanOperationType.optPaymentApproval.Value = 1
                                        Else: frmALISPLoanOperationType.optPaymentApproval.Value = 0
                                        End If
                            
                                        If !PaymentPreparation = True Then
                                                frmALISPLoanOperationType.optPaymentPreparation.Value = 1
                                        Else: frmALISPLoanOperationType.optPaymentPreparation.Value = 0
                                        End If
                            
                                        If !PaymentAuthorization = True Then
                                                frmALISPLoanOperationType.optPaymentAuthorization.Value = 1
                                        Else: frmALISPLoanOperationType.optPaymentAuthorization.Value = 0
                                        End If
                                        
                                        If !ChequeApproval = True Then
                                                frmALISPLoanOperationType.optchequeApproval.Value = 1
                                        Else: frmALISPLoanOperationType.optchequeApproval.Value = 0
                                        End If
                                        
                                        If !ChequePreparation = True Then
                                                frmALISPLoanOperationType.optChequePreparation.Value = 1
                                        Else: frmALISPLoanOperationType.optChequePreparation.Value = 0
                                        End If
                                        
                                        If !ChequeAuthorization = True Then
                                                frmALISPLoanOperationType.optChequeAuthorization.Value = 1
                                        Else: frmALISPLoanOperationType.optChequeAuthorization.Value = 0
                                        End If
                                        
                                        If !ChequeIssuance = True Then
                                                frmALISPLoanOperationType.optChequeIssuance.Value = 1
                                        Else: frmALISPLoanOperationType.optChequeIssuance.Value = 0
                                        End If

                                        Edit = True
                                End If
                        End With
        
                        If Edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ALISPLoanOperationType WHERE OperationType = '" & txtOperationType.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                                !OperationType = frmALISPLoanOperationType.txtOperationType
                                !Checked = frmALISPLoanOperationType.optChecked
                                !Approved = frmALISPLoanOperationType.optApproved
                                !authorized = frmALISPLoanOperationType.optAuthorized
                                !Description = frmALISPLoanOperationType.txtDescription
                                
                                If frmALISPLoanOperationType.optReinstatementPreparation = True Then
                                        !ReinstatementPreparation = 1
                                    Else: !ReinstatementPreparation = 0
                                End If
                                
                                If frmALISPLoanOperationType.optReinstatementApproval = True Then
                                        !ReinstatementApproval = 1
                                Else: !ReinstatementApproval = 0
                                End If
                                
                                If frmALISPLoanOperationType.optReinstatementAuthorization = True Then
                                        !ReinstatementAuthorization = 1
                                Else: !ReinstatementAuthorization = 0
                                End If

                                
                                
                                If frmALISPLoanOperationType.optClaimRegPreparation = True Then
                                        !RegistrationPreparation = 1
                                    Else: !RegistrationPreparation = 0
                                End If
                                
                                If frmALISPLoanOperationType.optClaimRegApproval = True Then
                                        !RegistrationApproval = 1
                                Else: !RegistrationApproval = 0
                                End If
                                
                                If frmALISPLoanOperationType.optClaimRegAuthorization = True Then
                                        !RegistrationAuthorization = 1
                                Else: !RegistrationAuthorization = 0
                                End If
                        
                                If frmALISPLoanOperationType.optDischargePreparation = True Then
                                        !DischargePreparation = 1
                                    Else: !DischargePreparation = 0
                                End If
                                
                                
                                If frmALISPLoanOperationType.optDischargeApproval = True Then
                                        !DischargeApproval = 1
                                Else: !DischargeApproval = 0
                                End If
                                
                                If frmALISPLoanOperationType.optDischargeAuthorization = True Then
                                        !DischargeAuthorization = 1
                                Else: !DischargeAuthorization = 0
                                End If
                                
                                If frmALISPLoanOperationType.optPaymentPreparation = True Then
                                    !PaymentPreparation = 1
                                Else: !PaymentPreparation = 0
                                End If
                                
                                
                                If frmALISPLoanOperationType.optPaymentApproval = True Then
                                        !PaymentApproval = 1
                                        Else: !PaymentApproval = 0
                                End If
                                
                                        
                                If frmALISPLoanOperationType.optPaymentAuthorization = True Then
                                        !PaymentAuthorization = 1
                                Else: !PaymentAuthorization = 0
                                End If
                                
                                If frmALISPLoanOperationType.optChequePreparation = True Then
                                        !ChequePreparation = 1
                                Else: !ChequePreparation = 0
                                End If
                                
                                If frmALISPLoanOperationType.optchequeApproval = True Then
                                        !ChequeApproval = 1
                                Else: !ChequeApproval = 0
                                End If
                                
                                If frmALISPLoanOperationType.optChequeAuthorization = True Then
                                        !ChequeAuthorization = 1
                                Else: !ChequeAuthorization = 0
                                End If
                                
                                If frmALISPLoanOperationType.optChequeIssuance = True Then
                                        !ChequeIssuance = 1
                                    Else: !ChequeIssuance = 0
                                End If
                            .Update
                            .Requery
                            Edit = False
                    End With
                
                    cmdEdit.Caption = "&Edit"
            Case Else
        
            Exit Sub

        End Select

Exit Sub

err:

    If err.Number = 40009 Then
            MsgBox "Record requested does not exist in the Database! Check your Entries.", vbInformation, "Searching."
                rsFind.Requery

            If rsFind.BOF Then Exit Sub
                rsFind.MoveFirst

    ElseIf err.Number = 3021 Then
            MsgBox "Requested record not found! Refresh the database and try the search again...or Check your entries.", vbInformation, "Searching."
                rsFind.Requery

            If rsFind.BOF Then Exit Sub
                rsFind.MoveFirst
    Else
                UpdateErrorMessage
End If

End Sub


Private Sub cmdFirstCode_Click(Index As Integer)
On Error GoTo err

        cmdUpdate.Enabled = False

        With rsSAVE
        If .EOF And .BOF Then Exit Sub
    
                    Select Case Index
                                Case 0
                                    .MoveFirst
                                Case 1
                                    .MovePrevious
                                    If .BOF Then .MoveFirst
                                Case 2
                                    .MoveNext
                                    If .EOF Then .MoveLast
                                Case 3
                                    .MoveLast
                    End Select
        End With

                    showRELN
                    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ValidateRELN()
On Error GoTo err

        bsaveRECORD = False
        
        If Screen.ActiveForm.txtOperationType.Text = "" Then
                MsgBox "The Operation Type MUST be Entered"
                Screen.ActiveForm.txtOperationType.SetFocus
        ElseIf Screen.ActiveForm.txtOperationType.Text <= "" Then
                MsgBox "The Description of the Operation cannot be Left Blank"
                txtOperationType.SetFocus
        Else
                bsaveRECORD = True
        End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveRecord()
On Error GoTo err
    Dim rsSAVE As ADODB.Recordset, strSQL1 As String
    Set rsSAVE = New ADODB.Recordset
    
    strSQL1 = "Select * from ALISPLoanOperationTyPe"
    rsSAVE.Open strSQL1, cnALIS, adOpenKeyset, adLockOptimistic
    
   With rsSAVE
        .AddNew
        !OperationType = frmALISPLoanOperationType.txtOperationType
        !Checked = frmALISPLoanOperationType.optChecked
        !Approved = frmALISPLoanOperationType.optApproved
        !authorized = frmALISPLoanOperationType.optAuthorized
        !Description = frmALISPLoanOperationType.txtDescription
        
        If frmALISPLoanOperationType.optMedicalPreparation = True Then
                !MedicalPreparation = 1
            Else: !MedicalPreparation = 0
        End If
        
        If frmALISPLoanOperationType.optMedicalApproval = True Then
                !MedicalApproval = 1
        Else: !MedicalApproval = 0
        End If
        
        If frmALISPLoanOperationType.optMedicalAuthorization = True Then
                !MedicalAuthorization = 1
        Else: !MedicalAuthorization = 0
        End If

        If frmALISPLoanOperationType.optProposalPreparation = True Then
                !ProposalPreparation = 1
            Else: !ProposalPreparation = 0
        End If
        
        If frmALISPLoanOperationType.optProposalApproval = True Then
                !ProposalApproval = 1
        Else: !ProposalApproval = 0
        End If
        
        If frmALISPLoanOperationType.optProposalAuthorization = True Then
                !ProposalAuthorization = 1
        Else: !ProposalAuthorization = 0
        End If

        
        If frmALISPLoanOperationType.optPolicyPreparation = True Then
                !PolicyPreparation = 1
            Else: !PolicyPreparation = 0
        End If
        
        If frmALISPLoanOperationType.optPolicyApproval = True Then
                !PolicyApproval = 1
        Else: !PolicyApproval = 0
        End If
        
        If frmALISPLoanOperationType.optPolicyAuthorization = True Then
                !PolicyAuthorization = 1
        Else: !PolicyAuthorization = 0
        End If

        
        
        If frmALISPLoanOperationType.optPaidupPreparation = True Then
                !paidupPreparation = 1
            Else: !paidupPreparation = 0
        End If
        
        If frmALISPLoanOperationType.optPaidupApproval = True Then
                !paidupApproval = 1
        Else: !paidupApproval = 0
        End If
        
        If frmALISPLoanOperationType.optPaidupAuthorization = True Then
                !paidupAuthorization = 1
        Else: !paidupAuthorization = 0
        End If

        
        
        
        If frmALISPLoanOperationType.optReinstatementPreparation = True Then
                !ReinstatementPreparation = 1
            Else: !ReinstatementPreparation = 0
        End If
        
        If frmALISPLoanOperationType.optReinstatementApproval = True Then
                !ReinstatementApproval = 1
        Else: !ReinstatementApproval = 0
        End If
        
        If frmALISPLoanOperationType.optReinstatementAuthorization = True Then
                !ReinstatementAuthorization = 1
        Else: !ReinstatementAuthorization = 0
        End If

        If frmALISPLoanOperationType.optReinstatementPreparation = True Then
                !RegistrationPreparation = 1
            Else: !RegistrationPreparation = 0
        End If
        
        If frmALISPLoanOperationType.optClaimRegApproval = True Then
                !RegistrationApproval = 1
        Else: !RegistrationApproval = 0
        End If
        
        If frmALISPLoanOperationType.optClaimRegAuthorization = True Then
                !RegistrationAuthorization = 1
        Else: !RegistrationAuthorization = 0
        End If

        If frmALISPLoanOperationType.optDischargePreparation = True Then
                !DischargePreparation = 1
            Else: !DischargePreparation = 0
        End If
        
        
        If frmALISPLoanOperationType.optDischargeApproval = True Then
                !DischargeApproval = 1
        Else: !DischargeApproval = 0
        End If
        
        If frmALISPLoanOperationType.optDischargeAuthorization = True Then
                !DischargeAuthorization = 1
        Else: !DischargeAuthorization = 0
        End If
        
        If frmALISPLoanOperationType.optPaymentPreparation = True Then
            !PaymentPreparation = 1
        Else: !PaymentPreparation = 0
        End If
        
        
        If frmALISPLoanOperationType.optPaymentApproval = True Then
                !PaymentApproval = 1
                Else: !PaymentApproval = 0
        End If
        
                
        If frmALISPLoanOperationType.optPaymentAuthorization = True Then
                !PaymentAuthorization = 1
        Else: !PaymentAuthorization = 0
        End If
        
        If frmALISPLoanOperationType.optChequePreparation = True Then
                !ChequePreparation = 1
        Else: !ChequePreparation = 0
        End If
        
        If frmALISPLoanOperationType.optchequeApproval = True Then
                !ChequeApproval = 1
        Else: !ChequeApproval = 0
        End If
        
        If frmALISPLoanOperationType.optChequeAuthorization = True Then
                !ChequeAuthorization = 1
        Else: !ChequeAuthorization = 0
        End If
        
        If frmALISPLoanOperationType.optChequeIssuance = True Then
                !ChequeIssuance = 1
            Else: !ChequeIssuance = 0
        End If
        bsaveRECORD = False
        
         .Update
         .Requery
  End With
Exit Sub

err:
    If err.Number = -2147217873 Then
    MsgBox "The record cannot be saved in the database because it could duplicate an existing record!.", vbOKOnly, "Canceling Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
            cmdCancel.SetFocus
            'cmdnew.Enabled = True
            SendKeys "{home} + {end}"

            If rsSAVE.EOF Then Exit Sub
            rsSAVE.MoveLast
    ElseIf err.Number = -2147467259 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
            'cmdnew.Enabled = True
            'txttranno.SetFocus

    ElseIf err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If


    
End Sub


Private Sub cmdUpdate_Click()
        bsaveRECORD = True
        ValidateRELN
        If bsaveRECORD = True Then
            SaveRecord
                If bsaveRECORD = False Then
                    enableButtons
                    clearALLRECORD
                    disableALLRECORD
                End If
        End If

        loadAPPROVERGRID
End Sub

Private Sub cmdSearch_Click()
On Error GoTo err

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        strQRE = InputBox("Enter The Operation type to search.", "Search Value")
        
        rsFind.Open "SELECT * FROM ALISPLoanOperationType WHERE OperationType = '" & strQRE & "';", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                    txtOperationType = !OperationType
                    optChecked = !Checked
                    optApproved = !Approved
                    txtDescription = !Description
                    optAuthorized = !authorized
                    
                    If !ReinstatementApproval = True Then
                            frmALISPLoanOperationType.optReinstatementApproval.Value = True
                    Else: frmALISPLoanOperationType.optReinstatementApproval.Value = False
                    End If
                    
                    If !ReinstatementPreparation = True Then
                            frmALISPLoanOperationType.optReinstatementPreparation.Value = True
                    Else: frmALISPLoanOperationType.optReinstatementPreparation.Value = False
                    End If
                    
                    If !ReinstatementAuthorization = True Then
                            frmALISPLoanOperationType.optReinstatementAuthorization.Value = True
                    Else: frmALISPLoanOperationType.optReinstatementAuthorization.Value = False
                    End If

                    
                    If !MedicalApproval = True Then
                            frmALISPLoanOperationType.optMedicalApproval.Value = True
                    Else: frmALISPLoanOperationType.optMedicalApproval.Value = False
                    End If
                    
                    If !MedicalPreparation = True Then
                            frmALISPLoanOperationType.optMedicalPreparation.Value = True
                    Else: frmALISPLoanOperationType.optMedicalPreparation.Value = False
                    End If
                    
                    If !MedicalAuthorization = True Then
                            frmALISPLoanOperationType.optMedicalAuthorization.Value = True
                    Else: frmALISPLoanOperationType.optMedicalAuthorization.Value = False
                    End If

                    
                    If !ProposalApproval = True Then
                            frmALISPLoanOperationType.optProposalApproval.Value = True
                    Else: frmALISPLoanOperationType.optProposalApproval.Value = False
                    End If
                    
                    If !ProposalPreparation = True Then
                            frmALISPLoanOperationType.optProposalPreparation.Value = True
                    Else: frmALISPLoanOperationType.optProposalPreparation.Value = False
                    End If
                    
                    If !ProposalAuthorization = True Then
                            frmALISPLoanOperationType.optProposalAuthorization.Value = True
                    Else: frmALISPLoanOperationType.optProposalAuthorization.Value = False
                    End If

                    
                    
                    If !PolicyApproval = True Then
                            frmALISPLoanOperationType.optPolicyApproval.Value = True
                    Else: frmALISPLoanOperationType.optPolicyApproval.Value = False
                    End If
                    
                    If !PolicyPreparation = True Then
                            frmALISPLoanOperationType.optPolicyPreparation.Value = True
                    Else: frmALISPLoanOperationType.optPolicyPreparation.Value = False
                    End If
                    
                    If !PolicyAuthorization = True Then
                            frmALISPLoanOperationType.optPolicyAuthorization.Value = True
                    Else: frmALISPLoanOperationType.optPolicyAuthorization.Value = False
                    End If

                    
                    
                    
                    If !paidupApproval = True Then
                            frmALISPLoanOperationType.optPaidupApproval.Value = True
                    Else: frmALISPLoanOperationType.optPaidupApproval.Value = False
                    End If
                    
                    If !paidupPreparation = True Then
                            frmALISPLoanOperationType.optPaidupPreparation.Value = True
                    Else: frmALISPLoanOperationType.optPaidupPreparation.Value = False
                    End If
                    
                    If !paidupAuthorization = True Then
                            frmALISPLoanOperationType.optPaidupAuthorization.Value = True
                    Else: frmALISPLoanOperationType.optPaidupAuthorization.Value = False
                    End If

                    If !DischargeApproval = True Then
                            frmALISPLoanOperationType.optDischargeApproval.Value = 1
                    Else: frmALISPLoanOperationType.optDischargeApproval.Value = 0
                    End If
                    
                    If !DischargePreparation = True Then
                            frmALISPLoanOperationType.optDischargePreparation.Value = 1
                    Else: frmALISPLoanOperationType.optDischargePreparation.Value = 0
                    End If
                    
                    If !DischargeAuthorization = True Then
                            frmALISPLoanOperationType.optDischargeAuthorization.Value = 1
                    Else: frmALISPLoanOperationType.optDischargeAuthorization.Value = 0
                    End If
            
                    If !PaymentApproval = True Then
                            frmALISPLoanOperationType.optPaymentApproval.Value = 1
                    Else: frmALISPLoanOperationType.optPaymentApproval.Value = 0
                    End If
        
                    If !PaymentPreparation = True Then
                            frmALISPLoanOperationType.optPaymentPreparation.Value = 1
                    Else: frmALISPLoanOperationType.optPaymentPreparation.Value = 0
                    End If
        
                    If !PaymentAuthorization = True Then
                            frmALISPLoanOperationType.optPaymentAuthorization.Value = 1
                    Else: frmALISPLoanOperationType.optPaymentAuthorization.Value = 0
                    End If
                    
                    If !ChequeApproval = True Then
                            frmALISPLoanOperationType.optchequeApproval.Value = 1
                    Else: frmALISPLoanOperationType.optchequeApproval.Value = 0
                    End If
                    
                    If !ChequePreparation = True Then
                            frmALISPLoanOperationType.optChequePreparation.Value = 1
                    Else: frmALISPLoanOperationType.optChequePreparation.Value = 0
                    End If
                    
                    If !ChequeAuthorization = True Then
                            frmALISPLoanOperationType.optChequeAuthorization.Value = 1
                    Else: frmALISPLoanOperationType.optChequeAuthorization.Value = 0
                    End If
                    
                    If !ChequeIssuance = True Then
                            frmALISPLoanOperationType.optChequeIssuance.Value = 1
                    Else: frmALISPLoanOperationType.optChequeIssuance.Value = 0
                    End If


                    Edit = True
                End If

            End With

        Exit Sub

err:
            ErrorMessage

End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableButtons
    loadAPPROVERGRID
End Sub

Private Sub Form_Load()

    OpenConnection
      
    Set rsSAVE = New Recordset
            strRELN = "SELECT * from ALISPLoanOperationType;"

    rsSAVE.Open strRELN, cnALIS, adOpenKeyset, adLockOptimistic

    'disableRELN

End Sub



