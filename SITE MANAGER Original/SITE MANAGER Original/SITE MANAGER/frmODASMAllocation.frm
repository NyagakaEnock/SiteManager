VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASMAllocation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lease Preparation"
   ClientHeight    =   8025
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13365
   Icon            =   "frmODASMAllocation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   13365
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   240
      TabIndex        =   37
      Top             =   3600
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Select Property"
      TabPicture(0)   =   "frmODASMAllocation.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListProperties"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Leasable Structures"
      TabPicture(1)   =   "frmODASMAllocation.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView3"
      Tab(1).Control(1)=   "chkLeaseAll"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Registered LandLords"
      TabPicture(2)   =   "frmODASMAllocation.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView1"
      Tab(2).Control(1)=   "txtSearchName"
      Tab(2).Control(2)=   "cmdSearch"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Properties Belonging to this LandLord"
      TabPicture(3)   =   "frmODASMAllocation.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Installment"
      TabPicture(4)   =   "frmODASMAllocation.frx":04B2
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "ListALLInstallments"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
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
         Left            =   -64320
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1860
         Width           =   1695
      End
      Begin VB.TextBox txtSearchName 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -64320
         TabIndex        =   43
         Top             =   1380
         Width           =   1695
      End
      Begin VB.CheckBox chkLeaseAll 
         Caption         =   "Lease All"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   420
         Width           =   1935
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   38
         Top             =   420
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6376
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   39
         Top             =   420
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6588
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   41
         Top             =   780
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   5741
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin MSComctlLib.ListView ListProperties 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   42
         Top             =   420
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   6376
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
      Begin MSComctlLib.ListView ListALLInstallments 
         Height          =   3615
         Left            =   120
         TabIndex        =   45
         Top             =   420
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6376
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   12735
      Begin VB.Frame Frame5 
         Caption         =   "Rent increment Status"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         TabIndex        =   32
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton OptVarrying 
            Caption         =   "Varrying"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1080
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OptStandard 
            Caption         =   "Standard"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.OptionButton optAmount 
         Caption         =   "Amt"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4200
         TabIndex        =   31
         Top             =   2430
         Width           =   615
      End
      Begin VB.OptionButton optPercentage 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3120
         TabIndex        =   30
         Top             =   2430
         Width           =   615
      End
      Begin VB.CheckBox chkYes 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtPercentage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         TabIndex        =   27
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808000&
         Caption         =   "Schedule"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtWitnessCoy 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6840
         TabIndex        =   22
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtSignedBy 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6840
         TabIndex        =   21
         Top             =   1920
         Width           =   3615
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00808000&
         Caption         =   "Installments"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   10920
         Picture         =   "frmODASMAllocation.frx":04CE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtpaymentMode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         TabIndex        =   17
         Top             =   1140
         Width           =   1695
      End
      Begin VB.ComboBox cboPaymentMode 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6840
         TabIndex        =   16
         Top             =   1133
         Width           =   1935
      End
      Begin VB.TextBox txtContractNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPickerAgreementDate 
         Height          =   330
         Left            =   3840
         TabIndex        =   12
         Top             =   1140
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         Format          =   15794177
         CurrentDate     =   38365
      End
      Begin VB.TextBox txtWitnessLandLord 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1530
         Width           =   3615
      End
      Begin VB.CheckBox chkDeallocate 
         Caption         =   "De Allocate?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPlotNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtAgreementDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtNames 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1530
         Width           =   2175
      End
      Begin VB.TextBox txtLandLordNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   750
         Width           =   1695
      End
      Begin VB.TextBox txtMastNo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   25
         Top             =   1530
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtInstallmentNo 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   35
         Text            =   " "
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "%/Amount"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   36
         Top             =   495
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "By"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Annual Rent Incmt."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Witness Coy"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   20
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Signed By"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Payment MODE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   15
         Top             =   1178
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Contract No"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Land Lord's Witness"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   1568
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Plot No"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Agreement Date"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   1178
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1568
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Land Lord No"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   750
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   7080
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
            Picture         =   "frmODASMAllocation.frx":0A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocation.frx":107A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocation.frx":14CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocation.frx":17E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocation.frx":1E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocation.frx":24DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocation.frx":292C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1(1)"
      DisabledImageList=   "ImageList1(1)"
      HotImageList    =   "ImageList1(1)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New &Record "
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
            Caption         =   "&Help"
            Key             =   "H"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   9720
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
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
               Picture         =   "frmODASMAllocation.frx":2FA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMAllocation.frx":3620
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMAllocation.frx":3B62
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMAllocation.frx":3FB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMAllocation.frx":42CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMAllocation.frx":4948
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMAllocation.frx":4FC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMAllocation.frx":5414
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear the &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnumm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show/View"
      Begin VB.Menu mnuClosedJobs 
         Caption         =   "Closed Jobs"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuKJHGFDGFVHJ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullInventory 
         Caption         =   "Full Inventory"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHow 
         Caption         =   "How to use this System"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmODASMAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsALLOCATION As clsODASAllocation, MyCommonData As clsCommonData

Private Sub cboPaymentMode_Click()
    Me.txtpaymentMode.SetFocus
End Sub

Private Sub cboPaymentMode_GotFocus()
    With Me
        If .cboPaymentMode.ListCount <> 0 Then Exit Sub
            .cboPaymentMode.Clear
            AttachSQL = "SELECT (ODASPPaymentMode.PaymentModeDescription)as selectfield,ODASPPaymentMode.* FROM ODASPPaymentMode ;"
            AttachDropDowns
    End With
End Sub

Private Sub cboPaymentMode_LostFocus()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentModeDescription = '" & .cboPaymentMode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
            .txtpaymentMode.Text = rsFindRecord!PaymentModeDescription
            .cboPaymentMode.Text = rsFindRecord!PaymentMode
        
    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub chkLeaseAll_Click()
On Error GoTo err
    With Me
        j = .ListView3.ListItems.Count
        If j = 0 Or .ListView3.View <> lvwReport Then
            .chkLeaseAll.Value = 0: Exit Sub
        Else
            If .chkLeaseAll.Value = 1 Then
                For i = 1 To j
                    .ListView3.ListItems(i).Checked = True
                Next i
                k = .ListView3.ListItems.Count
            ElseIf .chkLeaseAll.Value = 0 Then
                For i = 1 To j
                    .ListView3.ListItems(i).Checked = False
                Next i
                k = 0
            End If
        End If
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub chkYes_Click()
    If Me.chkYes.Value = 0 Then
        Me.txtPercentage = Empty: Me.txtPercentage.Locked = True: Me.optAmount.Value = False: Me.optPercentage.Value = False: Me.optAmount.Enabled = False: Me.optPercentage.Enabled = False
    Else
        Me.txtPercentage.Locked = False: Me.optAmount.Enabled = True: Me.optPercentage.Enabled = True
    End If
End Sub

Private Sub cmdPrint_Click()
On Error GoTo err
    If Me.txtContractNo.Text = Empty Then Exit Sub
        CurrentRecord = Me.txtContractNo.Text
        INPQRY2 = InputBox("Please enter the year for which you want to view the installments", "Yearly Installments")
        Load frmODASRRentPaymentInstallments
        frmODASRRentPaymentInstallments.Show vbModal
Exit Sub
err:
ErrorMessage
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

Private Sub Command1_Click()
On Error GoTo err
    If Me.txtContractNo = Empty Then Exit Sub
        CurrentRecord = Me.txtContractNo.Text
        Load frmODASRagreementSchedule
        frmODASRagreementSchedule.Show vbModal
Exit Sub
err:
ErrorMessage

End Sub

Private Sub DTPickerAgreementDate_CloseUp()
    With Me
        .txtAgreementDate.Text = .DTPickerAgreementDate.Value
    End With
End Sub

Private Sub Form_Activate()
    getLANDLORDS
    If bPlotRenewal = True Then
        getMastsToLease
    Else:
        getLeasableMasts
    End If
    disableFRAME
    If SchedulingMain.txtTask.Text = "K22" Then
        LeasedMasts
        loadMoreDETAILS
        frmODASMAllocation.Toolbar1.Buttons(2).Enabled = False
        frmODASMAllocation.cmdPrint.Enabled = False
        frmODASMAllocation.Command1.Enabled = False
    Else
        loadDEFAULTS
        LoadDEFAULT
    End If
    showALLLandLORDSites
End Sub

Public Sub getLANDLORDS()
On Error GoTo err
        With Me
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Land Lord No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Names", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 3
 
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                If bsearchRECORD = True Then
                        strSQL = "SELECT * FROM ODASPAccount Where CompanyName like '%" & Trim(.txtSearchName.Text) & "%' and Status = 'A' AND AccountType = 'LLORD' Order by AccountNo;"
                Else
                        strSQL = "SELECT * FROM ODASPAccount Where Status = 'A' AND AccountType = 'LLORD' oRDER BY AccountNo;"
                End If
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AccountNo))
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                                MyList.SubItems(2) = CStr(rsLIST!Status)
                        End If
                rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub setALLAcquiredSites()
On Error GoTo err
    
        With frmODASMAllocation
        
                .ListProperties.ListItems.Clear
                .ListProperties.ColumnHeaders.Clear
                
                .ListProperties.ColumnHeaders.Add , , "PlotNo", .ListProperties.Width / 3 ', lvwColumnCenter
                .ListProperties.ColumnHeaders.Add , , "PlotName", .ListProperties.Width / 3
                 .ListProperties.ColumnHeaders.Add , , "Physical Location", .ListProperties.Width / 3

                .ListProperties.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT *  FROM ODASPPlot  where (OnRoadReserve = 'N' or OnRoadReserve is null) ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                
                rsLIST.MoveFirst
                Do While rsLIST.EOF <> True
                
                Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open "SELECT *  FROM ODASPPlotMast where OwenedByClient = 'N' and  (LeasePrepared = 'N' or LeasePrepared is null) and PlotNo = '" & rsLIST!PlotNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsFindRecord.RecordCount > 0 Then
                
                Dim MyList As ListItem
                SchedulingMain.txtTotal.Text = rsLIST.RecordCount & "" & " Plots"
                            
                            Set MyList = .ListProperties.ListItems.Add(, , CStr(rsLIST!PlotNo))
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!PhysicalLocation) Then
                                MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation)
                            End If
                    End If
                rsLIST.MoveNext
            Loop
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Private Sub Form_Load()
    getLANDLORDS
    setALLAcquiredSites
End Sub
Private Sub loadDEFAULTS()
On Error GoTo err
    With frmODASMAllocation
        .txtAgreementDate.Text = Date
    End With

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

Exit Sub
err:
ErrorMessage

End Sub

Private Sub ListProperties_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListProperties.SortKey = ColumnHeader.Index - 1
    ListProperties.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListProperties_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
    Dim i, j As Double

    If Item.Checked = True Then
        
        j = Screen.ActiveForm.ListProperties.ListItems.Count
        
        If j = 0 Then Exit Sub
        
        For i = 1 To j
            If Screen.ActiveForm.ListProperties.ListItems(i) <> Item Then
               Screen.ActiveForm.ListProperties.ListItems(i).Checked = False
            End If
        Next i
        
        frmODASMAllocation.txtPlotNo.Text = Item.Text
        If bPlotRenewal = True Then
            getMastsToLease
        Else:
            getLeasableMasts
        End If
        disableFRAME
        If SchedulingMain.txtTask.Text = "K22" Then
            LeasedMasts
            loadMoreDETAILS
            frmODASMAllocation.Toolbar1.Buttons(2).Enabled = False
            frmODASMAllocation.cmdPrint.Enabled = False
            frmODASMAllocation.Command1.Enabled = False
        Else
            loadDEFAULTS
            LoadDEFAULT
        End If
        showALLLandLORDSites

    Else
        Item.Checked = False
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
        
        frmODASMAllocation.txtLandLordNo.Text = Item.Text
        frmODASMAllocation.txtNames.Text = Item.SubItems(1)
        showALLLandLORDSites

    Else
        Item.Checked = False
    End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
    Dim i, j As Double
    
    If Item.Checked = True Then
        
        j = Screen.ActiveForm.ListView2.ListItems.Count
        
        If j = 0 Then Exit Sub
        
        For i = 1 To j
            If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
               Screen.ActiveForm.ListView2.ListItems(i).Checked = False
            End If
        Next i
        
        frmODASMAllocation.txtContractNo.Text = Item.Text
                Set rsALLOCATION = New clsODASAllocation
                rsALLOCATION.loadRECORD
                Set rsALLOCATION = Nothing
        
        Else
        Item.Checked = False
    End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub disableFRAME()
On Error GoTo err
    
    With frmODASMAllocation
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub enableFRAME()
On Error GoTo err
    
    With frmODASMAllocation
        .Frame2.Enabled = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
With Me
    .txtMastNo.Text = Item
End With
End Sub

Private Sub optAmount_Click()
    If Me.optAmount.Value = True Then
        Me.txtPercentage.MaxLength = 5
    End If
End Sub

Private Sub Option2_Click()

End Sub

Private Sub optPercentage_Click()
    If Me.optPercentage.Value = True Then
        Me.txtPercentage.MaxLength = 3
    End If
End Sub

Private Sub OptStandard_Click()
  If Me.optPercentage.Value = True Then
        Me.txtInstallmentNo.Enabled = True
  End If
End Sub

Private Sub Optvarying_Click()
    If Me.optPercentage.Value = True Then
        Me.txtInstallmentNo.Enabled = True
         End If
      If MsgBox("Please Specify the Year", vbOK, "Entry Alert!!") = vbOK Then
         Exit Sub
      End If
 
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        
    With frmODASMAllocation
        Set rsALLOCATION = New clsODASAllocation

        Select Case Button.Key
        Case "N"
            Select Case Button.Caption
                    
                Case "New &Record "
                    If EditRecord Then Exit Sub
                    .ListView2.ListItems.Clear:
                    rsALLOCATION.enableRECORD
                    NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 5
                    enableFRAME
                    .cmdPrint.Enabled = False
                Case "&Save Record "
                        rsALLOCATION.updateRECORD
                Case Else
                    Exit Sub
                End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                    If NewRecord Then Exit Sub
                        If .txtContractNo.Text = Empty Then
                             MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                            .txtContractNo.SetFocus
                        Else
                             rsALLOCATION.enableRECORD
                            .txtContractNo.Locked = True
                             Button.Caption = "Save &Changes ": Button.Image = 5
                             EditRecord = True
                        End If
                Case "Save &Changes "
                    rsALLOCATION.updateRECORD
                    If ValidRecord = True Then
                        EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 5
                    End If
                Case Else
            End Select
        
        Case "S"
                Set rsFindRecord = New ADODB.Recordset
                INQUIRY = InputBox("Enter  the contract number to search and display...", "Search Values")
                rsFindRecord.Open "SELECT * FROM ODASMLeaseAgreement WHERE ContractNo = '" & INQUIRY & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsFindRecord.EOF And rsFindRecord.BOF Then
                    MsgBox "System could not match the requested record. Either it is deleted or currently missing", vbInformation + vbOKOnly + vbDefaultButton1, "Missing Records"
                Else
                    .txtContractNo.Text = rsFindRecord!ContractNo
                    .txtPlotNo.Text = rsFindRecord!PlotNo
                    .txtAgreementDate = rsFindRecord!AgreementDate
                    .txtLandLordNo.Text = rsFindRecord!LandLordNo
                    .txtWitnessLandLord.Text = rsFindRecord!WitnessLandLord
                    .txtWitnessCoy.Text = rsFindRecord!WitnessCoy
                    .txtSignedBy.Text = rsFindRecord!SignedBy
        
                        getLANDLORDS
                        loadDEFAULTS
                        disableFRAME
                        showALLLandLORDSites
                        showALLINSTALLMENTSDUE
                        
                strSQL = "select * from ODASPPlot Where PlotNo = '" & .txtPlotNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                
                    If rsSAVE!Status = "UN-ALLOCATED" Then
                        .chkDeallocate.Value = 1
                    Else: .chkDeallocate.Value = 0
                End If
                End If
        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
        Case "P"
                If frmODASMAllocation.txtContractNo.Text = Empty Then Exit Sub
                CurrentRecord = .txtContractNo.Text
                PrintDraft = True
                    k = 0: j = .ListView3.ListItems.Count
                    For i = 1 To j
                        If .ListView3.ListItems(i).Checked = True Then
                        k = k + 1
                        End If
                    Next i

                Load frmODASRContract1
                frmODASRContract1.Show 1, Me
        Case "H"
            .HelpCommonDialog.DialogTitle = "Using the Main System"
            .HelpCommonDialog.HelpFile = App.HelpFile
            .HelpCommonDialog.HelpContext = 15
            .HelpCommonDialog.HelpCommand = cdlHelpContext
            .HelpCommonDialog.ShowHelp
     
        Case Else
            Exit Sub
        End Select
        Set rsALLOCATION = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub txtPercentage_LostFocus()
        If Trim(Me.txtPercentage.Text) <= "" Then Exit Sub
        Me.txtPercentage.Text = FormatNumber(Me.txtPercentage.Text, 2)
End Sub
