VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASMSiteRegistrationOLD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Registration"
   ClientHeight    =   7995
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmODASMSiteRegistrationOLD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
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
            Picture         =   "frmODASMSiteRegistrationOLD.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRegistrationOLD.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRegistrationOLD.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRegistrationOLD.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRegistrationOLD.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRegistrationOLD.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRegistrationOLD.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
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
            Caption         =   "&Help System  "
            Key             =   "H"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10560
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   7335
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   11055
      Begin VB.TextBox txtZone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4440
         TabIndex        =   85
         Top             =   3840
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   6360
         TabIndex        =   81
         Top             =   6480
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.ComboBox cboCouncil 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   80
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox txtCouncil 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         TabIndex        =   79
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Frame Frame8 
         Caption         =   "Structure Ownership"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   73
         Top             =   120
         Width           =   3015
         Begin VB.OptionButton optEAGLE 
            Caption         =   "Firm/Company"
            Height          =   255
            Left            =   1560
            TabIndex        =   75
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optCLIENT 
            Caption         =   "Client"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "On Road Reserve?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   70
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton optNO 
            Caption         =   "No"
            Height          =   255
            Left            =   1800
            TabIndex        =   72
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton optYES 
            Caption         =   "Yes"
            Height          =   195
            Left            =   600
            TabIndex        =   71
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox txtOwnedBy 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   65
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Face Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   6120
         TabIndex        =   53
         Top             =   4080
         Width           =   4695
         Begin VB.CheckBox chkSiteActive 
            Caption         =   "Site Active?"
            Height          =   255
            Left            =   2880
            TabIndex        =   77
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtSiteDetails 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   720
            TabIndex        =   62
            Top             =   960
            Width           =   3855
         End
         Begin VB.Frame Frame6 
            Caption         =   "Sites Available"
            Height          =   1695
            Left            =   120
            TabIndex        =   60
            Top             =   1320
            Width           =   4455
            Begin MSComctlLib.ListView ListView3 
               Height          =   1335
               Left            =   120
               TabIndex        =   61
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   2355
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
         Begin VB.TextBox txtSiteAnnualRate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   720
            TabIndex        =   56
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtSiteStatus 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   2880
            TabIndex        =   55
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtSiteNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   720
            TabIndex        =   54
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Details"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Rate"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   660
            Width           =   345
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            Height          =   195
            Left            =   2280
            TabIndex        =   58
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Site No"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4440
         TabIndex        =   51
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtAnnualRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1680
         TabIndex        =   49
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Masts' Details "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   6120
         TabIndex        =   29
         Top             =   120
         Width           =   4695
         Begin VB.TextBox txtMeterNo 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   3360
            TabIndex        =   83
            Text            =   " "
            Top             =   2040
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTMastDate 
            Height          =   315
            Left            =   2160
            TabIndex        =   76
            Top             =   960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66584577
            CurrentDate     =   38425
         End
         Begin VB.ComboBox cboMediaSize 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1080
            TabIndex        =   68
            Top             =   2040
            Width           =   1335
         End
         Begin VB.ComboBox cboMedia 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1080
            TabIndex        =   66
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtMastLeaseDuration 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   3360
            TabIndex        =   46
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtMastStatus 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   3120
            TabIndex        =   44
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtMastExpiryDate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   3360
            TabIndex        =   42
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtMastDetails 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1080
            TabIndex        =   40
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtMastNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1080
            TabIndex        =   38
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtMastAnnualRate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   3360
            TabIndex        =   36
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtMastDOC 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1080
            TabIndex        =   34
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtMastAnnualRent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1080
            TabIndex        =   32
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Frame Frame4 
            Caption         =   "Mast Available"
            Height          =   1455
            Left            =   120
            TabIndex        =   30
            Top             =   2400
            Width           =   4455
            Begin MSComctlLib.ListView ListView2 
               Height          =   1095
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   1931
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
         Begin MSComCtl2.UpDown UpDownMastLeasePeriod 
            Height          =   315
            Left            =   4320
            TabIndex        =   48
            Top             =   960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label30 
            Caption         =   "MeterNo"
            Height          =   255
            Left            =   2520
            TabIndex        =   82
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "Size"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Media"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "M. LEASE"
            Height          =   195
            Left            =   2520
            TabIndex        =   47
            Top             =   1065
            Width           =   735
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            Height          =   195
            Left            =   2520
            TabIndex        =   45
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Expiry"
            Height          =   195
            Left            =   2520
            TabIndex        =   43
            Top             =   1740
            Width           =   420
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Details"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   630
            Width           =   480
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Mast No"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   300
            Width           =   600
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Rate"
            Height          =   195
            Left            =   2520
            TabIndex        =   37
            Top             =   1380
            Width           =   345
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "DOC"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1065
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Annual Rent"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1380
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "List All Plots"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   27
         Top             =   4680
         Width           =   5775
         Begin MSComctlLib.ListView ListView1 
            Height          =   2175
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   3836
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
      Begin MSComCtl2.UpDown UpDownLeasePeriod 
         Height          =   315
         Left            =   5640
         TabIndex        =   26
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPickerAcquisitionDate 
         Height          =   315
         Left            =   2880
         TabIndex        =   24
         Top             =   2760
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66584577
         CurrentDate     =   38298
      End
      Begin VB.TextBox txtExpiryDate 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtLeaseDuration 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4440
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtCommencementDate 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtTownCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtLRNo 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4440
         TabIndex        =   9
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtAnnualRent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtAcquiredBy 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtAcquisitionDate 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtPlotNo 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtPlotName 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtPhysicalAddress 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1680
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker DTPickerCommencementDate 
         Height          =   315
         Left            =   2880
         TabIndex        =   25
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66584577
         CurrentDate     =   38298
      End
      Begin VB.Label Label31 
         Caption         =   "Zone"
         Height          =   255
         Left            =   3480
         TabIndex        =   84
         Top             =   3870
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Council"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Acquired By"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   3480
         TabIndex        =   52
         Top             =   3540
         Width           =   450
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Annual Rate"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   3540
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date"
         Height          =   195
         Left            =   3480
         TabIndex        =   23
         Top             =   3180
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Lease Period"
         Height          =   195
         Left            =   3480
         TabIndex        =   22
         Top             =   2460
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Commencement Date"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2460
         Width           =   1530
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "LR No"
         Height          =   195
         Left            =   3480
         TabIndex        =   19
         Top             =   2820
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Annual Rent"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   3180
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Acquired By "
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2100
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AcquisitionDate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2805
         Width           =   1125
      End
      Begin VB.Label Label13 
         Caption         =   "Plot No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "PLOT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1680
         TabIndex        =   14
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Physical Address"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Town/City:"
         Height          =   195
         Left            =   3480
         TabIndex        =   12
         Top             =   2100
         Width           =   825
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
      Begin VB.Menu mnuCurrent 
         Caption         =   "Current Settings"
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
Attribute VB_Name = "frmODASMSiteRegistrationOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSITE As New clsODASSite, MyCommonData As clsCommonData

Private Sub cboCouncil_Click()
    Me.txtCouncil.SetFocus
End Sub

Private Sub cboCouncil_GotFocus()
On Error GoTo err
    With Me

        If .cboCouncil.ListCount <> 0 Then Exit Sub
        .cboCouncil.Clear
        AttachSQL = "SELECT  Council as SelectField , ODASPCouncil.* FROM ODASPCouncil WHERE TownCode = '" & .txtTownCode.Text & "';"
        AttachDropDowns
Exit Sub
err:
ErrorMessage
        
End With

End Sub

Private Sub cboCouncil_LostFocus()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPCouncil WHERE Council = '" & .cboCouncil.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.RecordCount = 0 Then Exit Sub
        
        .txtCouncil.Text = rsFindRecord!Council
        .cboCouncil.Text = rsFindRecord!CouncilCode
       
    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub cboMedia_Change()
With Me
        .cboMediaSize.Locked = False
        .cboMediaSize.Enabled = True
End With
End Sub

Private Sub cboMedia_Click()
    Me.cboMediaSize.SetFocus
End Sub

Private Sub cboMedia_GotFocus()
With Me
        If .cboMedia.ListCount <> 0 Then Exit Sub
        .cboMedia.Clear
        AttachSQL = "SELECT  MediaDescription as selectfield,ODASPMedia.* FROM ODASPMedia WHERE RequirePlot ='Y' ;"
        AttachDropDowns
End With

End Sub

Private Sub cboMedia_LostFocus()
On Error GoTo err
With Me
         Set rsFindRecord = New ADODB.Recordset
         rsFindRecord.Open "SELECT * FROM ODASPMedia WHERE MediaDescription ='" & .cboMedia.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
         .cboMedia.Text = rsFindRecord!MediaCode
End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub cboMediaSize_Click()
    Me.txtSiteDetails.SetFocus
End Sub
Private Sub cboMediaSize_GotFocus()
On Error GoTo err
    With Me

        If .cboMediaSize.ListCount <> 0 Then Exit Sub
        .cboMediaSize.Clear
        AttachSQL = "SELECT  MediaSize as SelectField ,ODASPMediaSize.* FROM ODASPMediaSize WHERE MediaCode = '" & .cboMedia.Text & "';"
        AttachDropDowns
Exit Sub
err:
ErrorMessage
        
End With
End Sub
Private Sub cboMediaSize_LostFocus()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPLandRate WHERE TownCode = '" & .txtTownCode.Text & "' and MediaCode = '" & .cboMedia.Text & "' and MediaSize  = '" & .cboMediaSize.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.RecordCount = 0 Then
        .txtSiteAnnualRate = 0
        Else
        .txtSiteAnnualRate.Text = rsFindRecord!Amount
        End If
    End With
Exit Sub
err:
ErrorMessage
End Sub
Private Sub DTMastDate_CloseUp()
    With Me
        .txtMastDOC.Text = .DTMastDate.Value
    End With
End Sub
Private Sub DTPickerAcquisitionDate_CloseUp()
On Error GoTo err
    With frmODASMSiteRegistration
            .txtAcquisitionDate.Text = .DTPickerAcquisitionDate.Value
    End With

Exit Sub

err:
    ErrorMessage

End Sub
Private Sub DTPickerCommencementDate_CloseUp()
On Error GoTo err
    With frmODASMSiteRegistration
            .txtCommencementDate.Text = .DTPickerCommencementDate.Value
            .txtAcquisitionDate.Text = .DTPickerCommencementDate.Value
    End With

Exit Sub

err:
    ErrorMessage

End Sub
Private Sub Form_Activate()
        disableALLRECORD
        Set MyCommonData = New clsCommonData
        Set rsSITE = New clsODASSite
        rsSITE.loadDEFAULTS
        Set rsSITE = Nothing
        showALLPLOTS
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

        If NewRecord Then
            Cancel = True
            MsgBox "System cannot close this form when registration of sites is in Process...", vbCritical + vbOKOnly, "Interruption of Job Brief creation process"
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
            
            frmODASMSiteRegistration.txtPlotNo.Text = Item.Text
            
            Set rsSITE = New clsODASSite
            rsSITE.loadRECORD
            showALLMASTS
            showALLSITES
            Set rsSITE = Nothing
            
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
            
            frmODASMSiteRegistration.txtMastNo.Text = Item.Text

            Set rsSITE = New clsODASSite
            rsSITE.loadMAST
            showALLSITES
            Set rsSITE = Nothing
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView3.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView3.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView3.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMSiteRegistration.txtSiteNo.Text = Item.Text

            Set rsSITE = New clsODASSite
            rsSITE.loadSITE
            Set rsSITE = Nothing
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Option1_Click()

End Sub
Private Sub optCLIENT_Click()
With Me
    .txtMastAnnualRent.BackColor = &HFFFFC0
    .txtMastAnnualRent.Locked = True
    .txtMastAnnualRent = 0
End With
End Sub
Private Sub optEAGLE_Click()
With Me
    If .optYES.Value = True Then Exit Sub
    .txtMastAnnualRent.BackColor = &HFFC0C0
    .txtMastAnnualRent.Locked = False
End With
End Sub
Private Sub optNo_Click()
With Me
    .txtPlotNo.Locked = False
    .txtPlotNo.BackColor = &HFFC0C0
    .txtPlotNo.Text = ""
    .txtMastAnnualRent.Locked = False
    .txtMastAnnualRent.BackColor = &HFFC0C0

End With
End Sub
Private Sub optYES_Click()
With Me
    .txtPlotNo.Locked = True
    .txtPlotNo.BackColor = &HFFFFC0
    .txtPlotNo.Text = GeneratePlotNumber
    .txtMastAnnualRent.Locked = True
    .txtMastAnnualRent.Text = 0
    .txtMastAnnualRent.BackColor = &HFFFFC0
End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        Set rsSITE = New clsODASSite
        
        With frmODASMSiteRegistration
        
        Select Case Button.Key
        Case "N"
            Select Case Button.Caption
                Case "New &Record "
                    
                    If EditRecord Then Exit Sub
                    rsSITE.enableSITE: rsSITE.enableMAST: rsSITE.enablePLOT: rsSITE.enableOptions
                    NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                    
                Case "&Save Record "
                    rsSITE.ValidDateRecord
                    If NewRecord = True Then
                    If bSaveRECORD = True Then
                            rsSITE.saveRecord
                            rsSITE.saveMAST
                            rsSITE.saveSITE
                            rsSITE.updateNoOfMasts
                            rsSITE.updateNoOfSites
                            rsSITE.updateALLSites
                            showALLMASTS
                            showALLPLOTS
                            showALLSITES
                            disableALLRECORD
                            NewRecord = False
                            
                            Button.Caption = "NEXT &MAST": Button.Image = 2
                            .Toolbar1.Buttons(3).Caption = "NEXT &FACE": .Toolbar1.Buttons(3).Image = 2
                            .Toolbar1.Buttons(4).Caption = "FINISH"
                    End If
                    End If

                 Case "NEXT &MAST"
                        rsSITE.clearSITE
                        rsSITE.enableSITE
                        rsSITE.clearMAST
                        .Frame8.Enabled = True
                        .optCLIENT.Enabled = True
                        .optEAGLE.Enabled = True
                        .optCLIENT.Value = False
                        .optEAGLE.Value = False
                        rsSITE.enableMAST
                        NewRecord = True: Button.Caption = "&Save Record ": bSaveRECORD = False: Button.Image = 4
                      
                Case Else
                    Exit Sub
                End Select
            
        Case "E"
                Select Case Button.Caption
                Case "&Edit/Change "
                    If NewRecord Then Exit Sub
                            If .txtPlotNo.Text = Empty Then
                                MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                               .txtPlotNo.SetFocus
                                Else
                               .txtPlotNo.Locked = True
                                rsSITE.enableSITE: rsSITE.enableMAST: rsSITE.enablePLOT: rsSITE.enableOptions
                                Button.Caption = "Save &Changes ": Button.Image = 4: rsSITE.enableSITE: rsSITE.enableMAST: rsSITE.enablePLOT
                                EditRecord = True
                            End If
                Case "NEXT &FACE"
                        rsSITE.clearSITE
                        rsSITE.enableSITE
                        NewRecord = True: .Toolbar1.Buttons(2).Image = 4: .Toolbar1.Buttons(2).Caption = "&Save Record ": bSaveRECORD = False
                
                Case "Save &Changes "
                    rsSITE.ValidDateRecord
                    If bSaveRECORD = True Then
                            rsSITE.saveRecord
                            rsSITE.saveMAST
                            rsSITE.saveSITE
                            rsSITE.updateANNUALRate
                            rsSITE.updateANNUALRent
                            rsSITE.updateNoOfMasts
                            rsSITE.updateNoOfSites
                            rsSITE.updateALLSites
                            If .optYES = True Then
                            End If
                            showALLMASTS
                            showALLPLOTS
                            showALLSITES
                            disableALLRECORD
                            NewRecord = False
                    
                    .Toolbar1.Buttons(2).Caption = "NEXT &MAST"
                    .Toolbar1.Buttons(3).Caption = "NEXT &FACE"
                    .Toolbar1.Buttons(4).Caption = "FINISH"
                 End If
                 Case "FINISH"
                            
                            .Toolbar1.Buttons(2).Caption = "New &Record "
                            .Toolbar1.Buttons(2).Image = 2
                            .Toolbar1.Buttons(3).Image = 5
                            .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                            NewRecord = False: EditRecord = False
                Case Else
                   
                    Exit Sub
                End Select
        Case "S"
                Select Case Button.Caption
                Case "FINISH"
                     
                            .Toolbar1.Buttons(2).Caption = "New &Record "
                            .Toolbar1.Buttons(2).Image = 2
                            .Toolbar1.Buttons(3).Image = 5
                            .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                            Button.Caption = "&Search/Find "
               
                Case "&Search/Find "
                            enableALLRECORD
                            .Toolbar1.Buttons(2).Caption = "Save &Changes "
                            .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                            Button.Caption = "Delete"
                    Case "Cancel"
                        cancelCMD
                    Case "Delete"
                        cmdDelete_Click
                        Button.Caption = "&Search/Find "
            Case Else
                
                    Exit Sub
                End Select

        Case "R"
                If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 2
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 5
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
        Case "P"
                Load frmODASRPlotSites
                frmODASRPlotSites.Show 1, frmODASMSiteRegistration
        Case "H"
            .HelpCommonDialog.DialogTitle = "Using the Main System"
            .HelpCommonDialog.HelpFile = App.HelpFile
            .HelpCommonDialog.HelpContext = 17
            .HelpCommonDialog.HelpCommand = cdlHelpContext
            .HelpCommonDialog.ShowHelp

             
        Case Else
            Exit Sub
        End Select
        End With

Set rsSITE = Nothing
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub loadExistingRECORD()
      With frmODASMSiteRegistration
                    INQUIRY = InputBox("Enter the Plot Number to search and display...", "Search Value request")
                    If Len(INQUIRY) = 0 Then Exit Sub
                    strSQL = "select * from ODASPPlot Where ODASPPlot.PlotNo = '" & INQUIRY & "' ;"
                    Set rsFIND = New ADODB.Recordset
                    rsFIND.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    If rsFIND.RecordCount = 0 Then
                        MsgBox "Requested Plot could not be found. Either it was deleted or you entered a wrong Number", vbCritical + vbOKOnly, "Records Search Engine"
                    Else
                           .txtPlotNo.Text = rsFIND!PlotNo
                           .txtAcquiredBy.Text = rsFIND!CreatedBy
                           .txtStatus.Text = rsFIND!Status
                           .txtTownCode.Text = rsFIND!TownCode
                           .txtLeaseDuration.Text = rsFIND!LeaseDuration
                           .txtLRNo.Text = rsFIND!LRNo
                           .txtPhysicalAddress.Text = rsFIND!PhysicalLocation
                           .txtCommencementDate.Text = rsFIND!CommencementDate
                    
                    Set rsFIND = Nothing
                    strSQL = Empty
                            showALLMASTS
                            showALLSITES
                            .Toolbar1.Buttons(2).Caption = "NEXT &MAST"
                            .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                            Button.Caption = "FINISH"
                    End If
                    
           End With

End Sub

Private Sub updatePLOTSITE()
On Error GoTo err
On Error Resume Next
        Me.ProgressBar2.Visible = True
        
          Dim rsPLOT As ADODB.Recordset, strPLOT As String
          Set rsPLOT = New ADODB.Recordset
        
          strPLOT = "SELECT * FROM ODASPPlotMast where MastNo = '" & Me.txtMastNo & "' ;"
          rsPLOT.Open strPLOT, cnCOMMON, adOpenKeyset, adLockOptimistic
                   
          If rsPLOT.EOF And rsPLOT.BOF Then Exit Sub
          
            
            rsPLOT.MoveFirst
            
            Do While Not rsPLOT.EOF
                  Set rsSAVE = New Recordset
                  strSQL = "SELECT * FROM ODASPPlotSite where MastNo='" & rsPLOT!MastNo & "';"
                  rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                  With rsSAVE
                  
                    .MoveFirst
                          Do While Not .EOF
                             If Me.optYES.Value = False Then
                                !ApprovedBy = CurrentUserName
                                !DateApproved = Date
                                !Approved = "Y"
                                
                                !AuthorizedBy = CurrentUserName
                                !DateAuthorized = Date
                                        
                                !Authorized = "Y"
                              End If
                                !Status = "SITE-AVAILABLE"
                                                
                                Dim SDate, EDate As Date
                                        SDate = rsPLOT!CommencementDate
                                        EDate = DateAdd("yyyy", rsPLOT!LeaseDuration, rsPLOT!CommencementDate)
                                
                                        Me.ProgressBar2.Max = DateDiff("d", SDate, EDate) + 1
                                        Me.ProgressBar2.Min = 0
                                        
                                        Do While SDate <= EDate
                                            Set rsSiteSchedule = New ADODB.Recordset
                                            rsSiteSchedule.Open "SELECT * FROM ODASMSiteSchedule WHERE SiteNo = '" & rsSAVE!SiteNo & "' and ScheduleDate = '" & Format(SDate, "MMMM dd,yyyy") & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                                            If rsSiteSchedule.RecordCount = 0 Then
                                                rsSiteSchedule.AddNew
                                                rsSiteSchedule!SiteNo = rsSAVE!SiteNo
                                                rsSiteSchedule!ScheduleDAte = SDate
                                                rsSiteSchedule!Reserved = "N"
                                                rsSiteSchedule!Allocated = "N"
                                            End If
                                            rsSiteSchedule.Update
                                            SDate = DateAdd("d", 1, SDate)
                                           Me.ProgressBar2.Value = Me.ProgressBar2.Value + 1
                                        Loop

                        .Update
                        Me.ProgressBar2.Value = 0
                       .MoveNext
                    Loop
                    End With
                rsPLOT.MoveNext
              Loop
            Me.ProgressBar2.Visible = False
       Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If
End Sub
Private Sub updatePLOTMAST()
On Error GoTo err
          Me.ProgressBar2.Visible = True
          Dim rsPLOT As ADODB.Recordset, strPLOT As String
          Set rsPLOT = New ADODB.Recordset
        
          strPLOT = "SELECT * FROM ODASPPlotMast where MastNo = '" & Me.txtMastNo & "' ;"
          rsPLOT.Open strPLOT, cnCOMMON, adOpenKeyset, adLockOptimistic
                   
          If rsPLOT.EOF Or rsPLOT.BOF Then Exit Sub
          
          rsPLOT.MoveFirst
          Do While Not rsPLOT.EOF

                  Set rsSAVE = New Recordset
                  strSQL = "SELECT * FROM ODASPPlotMast where MastNo = '" & rsPLOT!MastNo & "' ;"
                  rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                  With rsPLOT
                     If Me.optYES.Value = True Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        
                        !Approved = "Y"
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date
                        
                        !Authorized = "Y"
                      End If
                        !Status = "SITE-AVAILABLE"
                        !RentDueDate = Date
                        !RentDue = 0
                                
                        Dim SDate, EDate As Date
                        SDate = rsSAVE!CommencementDate
                        EDate = rsSAVE!expirydate
                
                        Me.ProgressBar2.Max = DateDiff("d", SDate, EDate) + 1
                        Me.ProgressBar2.Min = 0
                        
                        Do While SDate <= EDate
                            Set rsBillBoardSchedule = New ADODB.Recordset
                            rsBillBoardSchedule.Open "SELECT * FROM ODASMBillBoardSchedule WHERE MastNo = '" & rsSAVE!MastNo & "' and ScheduleDate = '" & Format(SDate, "MMMM dd,yyyy") & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                            If rsBillBoardSchedule.RecordCount = 0 Then
                                rsBillBoardSchedule.AddNew
                                rsBillBoardSchedule!MastNo = rsSAVE!MastNo
                                rsBillBoardSchedule!ScheduleDAte = SDate
                                rsBillBoardSchedule!Reserved = "N"
                                rsBillBoardSchedule!Allocated = "N"
                            End If
                            rsBillBoardSchedule.Update
                            SDate = DateAdd("d", 1, SDate)
                           Me.ProgressBar2.Value = Me.ProgressBar2.Value + 1
                        Loop

                        .Update
                        Me.ProgressBar2.Value = 0
                    End With
                    
                    rsPLOT.MoveNext
          Loop
          Me.ProgressBar2.Visible = False
Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If


End Sub
Private Function GeneratePlotNumber()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPPlot WHERE OnRoadReserve = 'Y' and TownCode = '" & .txtTownCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.RecordCount = 0 Then
            GeneratePlotNumber = .txtTownCode.Text & "/RR/" & "001"
        Else
            GeneratePlotNumber = .txtTownCode.Text & "/RR/" & rsFindRecord.RecordCount + 1
        End If
    End With
Exit Function
err:
ErrorMessage
End Function

Private Sub txtPhysicalAddress_LostFocus()
With Me
    .txtMastDetails.Text = .txtPhysicalAddress.Text & ""
End With
End Sub

Private Sub txtPlotNo_LostFocus()
With Me
    .txtLRNo.Text = .txtPlotNo.Text & ""
End With
End Sub

Private Sub txtSiteNo_LostFocus()
    showALLMASTS
End Sub
Private Sub UpDownLeasePeriod_Change()
On Error GoTo err
        With frmODASMSiteRegistration
            .txtLeaseDuration.Text = .UpDownLeasePeriod.Value
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub UpDownMastLeasePeriod_Change()
With Me
    .txtMastLeaseDuration.Text = .UpDownMastLeasePeriod.Value
    .txtMastExpiryDate.Text = DateAdd("YYYY", .UpDownMastLeasePeriod.Value, Format(.txtMastDOC.Text, "MMMM dd,yyyy"))
End With
End Sub
Private Sub cmdDelete_Click()
On Error GoTo err
With Me
If .txtPlotNo.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            Set rsCONTROL = New ADODB.Recordset
    
            strSQL = "Select * from ODASPPlot P,ODASPPlotMast PM,ODASPPlotSite PS Where (P.PlotNo = '" & Me.txtPlotNo.Text & "' or PM.PlotNo = '" & Me.txtPlotNo.Text & "' or PS.PlotNo = '" & Me.txtPlotNo.Text & "')"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            
            With rsCONTROL
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                'clearALLRECORD
               showALLPLOTS
               showALLMASTS
               showALLSITES
            End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
End With
Exit Sub

err:
    ErrorMessage

End Sub

