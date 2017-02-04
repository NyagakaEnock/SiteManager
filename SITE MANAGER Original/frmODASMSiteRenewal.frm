VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMSiteRenewal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Renewal"
   ClientHeight    =   8910
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13665
   Icon            =   "frmODASMSiteRenewal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   13335
      Begin VB.TextBox txtComments 
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
         Height          =   450
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   85
         Top             =   3120
         Width           =   10935
      End
      Begin VB.CheckBox chkWithLease 
         Caption         =   "With Lease"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtTownCity 
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
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   1260
         Width           =   2055
      End
      Begin VB.TextBox txtemailAddress 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   7200
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1260
         Width           =   2535
      End
      Begin VB.TextBox txtPostalAddress 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   7200
         MaxLength       =   30
         TabIndex        =   7
         Top             =   810
         Width           =   2535
      End
      Begin VB.TextBox txtMobileNo 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   11160
         MaxLength       =   15
         TabIndex        =   8
         Top             =   810
         Width           =   2055
      End
      Begin VB.TextBox txtLandLordName 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   9720
         MaxLength       =   50
         TabIndex        =   4
         Top             =   360
         Width           =   3495
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
         Left            =   7200
         TabIndex        =   75
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton optEAGLE 
         Caption         =   "Owned By Firm/Company"
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
         Left            =   7800
         TabIndex        =   17
         Top             =   2678
         Width           =   2055
      End
      Begin VB.OptionButton optCLIENT 
         Caption         =   "Owned by Client"
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
         Left            =   6240
         TabIndex        =   16
         Top             =   2678
         Width           =   1575
      End
      Begin VB.TextBox txtZone 
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
         Left            =   1200
         TabIndex        =   18
         Top             =   2640
         Width           =   4695
      End
      Begin VB.ComboBox cboCouncil 
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
         Left            =   7200
         TabIndex        =   13
         Top             =   1703
         Width           =   2535
      End
      Begin VB.TextBox txtCouncil 
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
         Left            =   9720
         TabIndex        =   27
         Top             =   1710
         Width           =   3495
      End
      Begin VB.TextBox txtStatus 
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
         Left            =   11160
         TabIndex        =   26
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtAnnualRate 
         Alignment       =   1  'Right Justify
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
         Left            =   11160
         TabIndex        =   25
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtExpiryDate 
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtLeaseDuration 
         Alignment       =   2  'Center
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
         TabIndex        =   23
         Top             =   1710
         Width           =   495
      End
      Begin VB.TextBox txtCommencementDate 
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
         Left            =   1200
         TabIndex        =   22
         Top             =   1710
         Width           =   1575
      End
      Begin VB.TextBox txtLRNo 
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
         TabIndex        =   21
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox txtAnnualRent 
         Alignment       =   1  'Right Justify
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
         Left            =   7200
         TabIndex        =   20
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtAcquiredBy 
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox txtAcquisitionDate 
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
         Left            =   1200
         TabIndex        =   19
         Top             =   2160
         Width           =   1575
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
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtPhysicalAddress 
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
         Left            =   1200
         TabIndex        =   6
         Top             =   810
         Width           =   4695
      End
      Begin VB.Frame Frame7 
         Caption         =   "On Road Reserve?"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   120
         Width           =   2775
         Begin VB.OptionButton optNO 
            Caption         =   "No"
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
            Left            =   1680
            TabIndex        =   3
            Top             =   210
            Width           =   615
         End
         Begin VB.OptionButton optYES 
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
            Height          =   195
            Left            =   600
            TabIndex        =   2
            Top             =   240
            Width           =   855
         End
      End
      Begin MSComCtl2.UpDown UpDownLeasePeriod 
         Height          =   315
         Left            =   4560
         TabIndex        =   12
         Top             =   1725
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPickerAcquisitionDate 
         Height          =   315
         Left            =   2760
         TabIndex        =   15
         Top             =   2168
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   105578497
         CurrentDate     =   38298
      End
      Begin MSComCtl2.DTPicker DTPickerCommencementDate 
         Height          =   315
         Left            =   2760
         TabIndex        =   11
         Top             =   1718
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   105578497
         CurrentDate     =   38298
      End
      Begin VB.Label Label51 
         Caption         =   "Comments:"
         Height          =   255
         Left            =   1440
         TabIndex        =   84
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Town:"
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
         Left            =   10080
         TabIndex        =   80
         Top             =   1313
         Width           =   405
      End
      Begin VB.Label Label21 
         Caption         =   "email Address"
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
         Left            =   6240
         TabIndex        =   78
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Postal Address"
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
         Left            =   6240
         TabIndex        =   77
         Top             =   848
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Mobile No"
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
         Left            =   10080
         TabIndex        =   76
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "Highway Zone"
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
         Left            =   120
         TabIndex        =   72
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Annual Rate "
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
         Left            =   10080
         TabIndex        =   71
         Top             =   2220
         Width           =   795
      End
      Begin VB.Label Label28 
         Caption         =   "Council"
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
         Left            =   6240
         TabIndex        =   39
         Top             =   1755
         Width           =   735
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   10320
         TabIndex        =   38
         Top             =   2693
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date"
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
         Left            =   3240
         TabIndex        =   37
         Top             =   2213
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Lease Period"
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
         Left            =   3240
         TabIndex        =   36
         Top             =   1763
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Start  Date"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1763
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "LR No"
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
         TabIndex        =   34
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Annual Rent"
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
         Left            =   6240
         TabIndex        =   33
         Top             =   2220
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Acquired By "
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
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AcquisitionDate"
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
         Left            =   120
         TabIndex        =   31
         Top             =   2213
         Width           =   930
      End
      Begin VB.Label Label13 
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
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Physical Location"
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
         Left            =   120
         TabIndex        =   29
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Land Lord"
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
         Left            =   6240
         TabIndex        =   28
         Top             =   390
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
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
            Picture         =   "frmODASMSiteRenewal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRenewal.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRenewal.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRenewal.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRenewal.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRenewal.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRenewal.frx":236E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteRenewal.frx":29E8
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
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   1164
      ButtonWidth     =   3307
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            Caption         =   "&Delete"
            Key             =   "D"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help System  "
            Key             =   "H"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10200
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   40
      Top             =   4440
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
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
      TabCaption(0)   =   "Mast Management"
      TabPicture(0)   =   "frmODASMSiteRenewal.frx":2DBA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Face Mangement"
      TabPicture(1)   =   "frmODASMSiteRenewal.frx":2DD6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "View All Plots"
      TabPicture(2)   =   "frmODASMSiteRenewal.frx":2DF2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Assign Specification"
      TabPicture(3)   =   "frmODASMSiteRenewal.frx":2E0E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame12"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame11"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame8"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "View Masts"
      TabPicture(4)   =   "frmODASMSiteRenewal.frx":2E2A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Select Landlord"
      TabPicture(5)   =   "frmODASMSiteRenewal.frx":2E46
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ListALLLandLords"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Contracts"
      TabPicture(6)   =   "frmODASMSiteRenewal.frx":2E62
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ListView4"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Installments"
      TabPicture(7)   =   "frmODASMSiteRenewal.frx":2E7E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label43"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Label44"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Label42"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Label41"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "txtSerialNo"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "Label40"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "Label26"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "Label19"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "ListALLInstallments"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "txtPaymentDueDate"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "txtInvoiceNo"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).Control(11)=   "cmdSave"
      Tab(7).Control(11).Enabled=   0   'False
      Tab(7).Control(12)=   "chkPaymentFlag"
      Tab(7).Control(12).Enabled=   0   'False
      Tab(7).Control(13)=   "txtContractYear"
      Tab(7).Control(13).Enabled=   0   'False
      Tab(7).Control(14)=   "txtCurrentPeriod"
      Tab(7).Control(14).Enabled=   0   'False
      Tab(7).Control(15)=   "txtTransactionNo"
      Tab(7).Control(15).Enabled=   0   'False
      Tab(7).Control(16)=   "txtInstallment"
      Tab(7).Control(16).Enabled=   0   'False
      Tab(7).Control(17)=   "txtPaymentDue"
      Tab(7).Control(17).Enabled=   0   'False
      Tab(7).Control(18)=   "txtInstallmentPercent"
      Tab(7).Control(18).Enabled=   0   'False
      Tab(7).ControlCount=   19
      Begin VB.Frame Frame2 
         Caption         =   "Masts' Details "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   120
         TabIndex        =   126
         Top             =   360
         Width           =   6495
         Begin VB.TextBox txtMastAnnualRent 
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
            Left            =   1440
            TabIndex        =   136
            Top             =   1305
            Width           =   1815
         End
         Begin VB.TextBox txtMastAnnualRate 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            TabIndex        =   135
            Top             =   1305
            Width           =   1935
         End
         Begin VB.TextBox txtMastNo 
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
            Left            =   1440
            TabIndex        =   134
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtMastDetails 
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
            Left            =   1440
            TabIndex        =   133
            Top             =   870
            Width           =   4695
         End
         Begin VB.TextBox txtMeterNo 
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
            Left            =   1440
            TabIndex        =   132
            Text            =   " "
            Top             =   1695
            Width           =   1815
         End
         Begin VB.ComboBox cboMediaSize 
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
            Left            =   1440
            TabIndex        =   131
            Top             =   2460
            Width           =   4695
         End
         Begin VB.ComboBox cboMedia 
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
            Left            =   1440
            TabIndex        =   130
            Top             =   2055
            Width           =   4695
         End
         Begin VB.TextBox txtTownCode 
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
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   129
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtNoofSites 
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
            Left            =   1440
            TabIndex        =   128
            Text            =   " "
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox txtNoofMasts 
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
            Left            =   2760
            TabIndex        =   127
            Text            =   " "
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label30 
            Caption         =   "Meter No"
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
            Left            =   480
            TabIndex        =   146
            Top             =   1740
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Annual Rent"
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
            Left            =   480
            TabIndex        =   145
            Top             =   1365
            Width           =   750
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Rate"
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
            Left            =   3480
            TabIndex        =   144
            Top             =   1365
            Width           =   285
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Mast No"
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
            Left            =   480
            TabIndex        =   143
            Top             =   540
            Width           =   525
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Mast Details"
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
            Left            =   480
            TabIndex        =   142
            Top             =   930
            Width           =   750
         End
         Begin VB.Label Label27 
            Caption         =   "Size"
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
            Left            =   480
            TabIndex        =   141
            Top             =   2505
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Media"
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
            Left            =   480
            TabIndex        =   140
            Top             =   2100
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Town:"
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
            Left            =   3360
            TabIndex        =   139
            Top             =   540
            Width           =   405
         End
         Begin VB.Label Label47 
            Caption         =   "No of Faces"
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
            Left            =   480
            TabIndex        =   138
            Top             =   2925
            Width           =   855
         End
         Begin VB.Label Label48 
            Caption         =   "No of Masts"
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
            Left            =   1920
            TabIndex        =   137
            Top             =   2925
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contract Details"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   6720
         TabIndex        =   106
         Top             =   360
         Width           =   6615
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
            Left            =   1320
            TabIndex        =   118
            Top             =   1020
            Width           =   2295
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
            Left            =   3600
            TabIndex        =   117
            Top             =   1020
            Width           =   2895
         End
         Begin VB.TextBox txtAnnualRentIncrement 
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
            Left            =   2640
            TabIndex        =   116
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtAnnualRentIncrementType 
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
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   115
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtRentVariationType 
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
            Left            =   6120
            TabIndex        =   114
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtIncrementStartYear 
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
            Left            =   6120
            TabIndex        =   113
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtIncrementFrequency 
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
            Left            =   6120
            TabIndex        =   112
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtPaymentInterval 
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
            Left            =   1680
            TabIndex        =   111
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdNewContract 
            Caption         =   "Renew Contract"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   110
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveContract 
            Caption         =   "Save Contract"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   109
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton cmdcancelContract 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   108
            Top             =   1440
            Width           =   1215
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
            Left            =   1320
            TabIndex        =   107
            Top             =   1462
            Width           =   1335
         End
         Begin VB.Label Label39 
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
            Left            =   240
            TabIndex        =   125
            Top             =   1065
            Width           =   1095
         End
         Begin VB.Label Label32 
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
            Left            =   1800
            TabIndex        =   124
            Top             =   285
            Width           =   855
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Increment Type"
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
            Left            =   240
            TabIndex        =   123
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Increment Starts From Installment No:"
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
            Left            =   3840
            TabIndex        =   122
            Top             =   180
            Width           =   2325
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Increment Interval"
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
            Left            =   3840
            TabIndex        =   121
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Payment  Interval"
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
            Left            =   240
            TabIndex        =   120
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label52 
            Caption         =   "Contract No:"
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
            Left            =   240
            TabIndex        =   119
            Top             =   1500
            Width           =   1095
         End
      End
      Begin VB.TextBox txtInstallmentPercent 
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
         Left            =   -70800
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   930
         Width           =   1455
      End
      Begin VB.TextBox txtPaymentDue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   -73680
         MaxLength       =   30
         TabIndex        =   94
         Top             =   930
         Width           =   1335
      End
      Begin VB.TextBox txtInstallment 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   -73680
         MaxLength       =   30
         TabIndex        =   93
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTransactionNo 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   -70800
         MaxLength       =   15
         TabIndex        =   92
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentPeriod 
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
         Left            =   -67680
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   930
         Width           =   1335
      End
      Begin VB.TextBox txtContractYear 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   -67680
         MaxLength       =   15
         TabIndex        =   90
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkPaymentFlag 
         Caption         =   "Payment Flag?"
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
         Left            =   -63360
         TabIndex        =   89
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0C000&
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70200
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtInvoiceNo 
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
         Left            =   -65040
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   930
         Width           =   1335
      End
      Begin VB.TextBox txtPaymentDueDate 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   -65040
         MaxLength       =   15
         TabIndex        =   86
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Face Details"
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
         Left            =   6720
         TabIndex        =   60
         Top             =   2340
         Width           =   6615
         Begin VB.TextBox txtSiteNo 
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
            Left            =   840
            TabIndex        =   64
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtSiteAnnualRate 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            TabIndex        =   63
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtSiteDetails 
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
            Left            =   840
            TabIndex        =   62
            Top             =   1035
            Width           =   5655
         End
         Begin VB.CheckBox chkSiteActive 
            Caption         =   "Site Active?"
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
            Left            =   4080
            TabIndex        =   61
            Top             =   668
            Width           =   1455
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   255
            Left            =   840
            TabIndex        =   68
            Top             =   1440
            Visible         =   0   'False
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Site No"
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
            Left            =   240
            TabIndex        =   67
            Top             =   293
            Width           =   450
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Rate"
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
            Left            =   3480
            TabIndex        =   66
            Top             =   293
            Width           =   285
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Details"
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
            Left            =   240
            TabIndex        =   65
            Top             =   1095
            Width           =   405
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3735
         Left            =   -69960
         TabIndex        =   45
         Top             =   420
         Width           =   2775
         Begin VB.CommandButton cmdPropertyDelete 
            BackColor       =   &H00808000&
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   2880
            Width           =   2535
         End
         Begin VB.CommandButton cmdPropertyADD 
            BackColor       =   &H00808000&
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Add New"
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtPropertyCode 
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
            Height          =   315
            Left            =   1320
            TabIndex        =   51
            Top             =   150
            Width           =   1335
         End
         Begin VB.TextBox txtPropertyDateAssigned 
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
            Left            =   1320
            TabIndex        =   50
            Top             =   525
            Width           =   1095
         End
         Begin VB.TextBox txtPropertyAmountDue 
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
            Left            =   1320
            TabIndex        =   49
            Top             =   915
            Width           =   1335
         End
         Begin VB.TextBox txtPropertyCommencementDate 
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
            Left            =   1320
            TabIndex        =   48
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtPropertyOtherDetails 
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
            Left            =   1320
            TabIndex        =   47
            Top             =   2190
            Width           =   1335
         End
         Begin VB.TextBox txtPropertyTransactionNo 
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
            Height          =   315
            Left            =   1320
            TabIndex        =   46
            Top             =   2520
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPickerPropertyCommencementDate 
            Height          =   315
            Left            =   2400
            TabIndex        =   52
            Top             =   1800
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   105578497
            CurrentDate     =   38300
         End
         Begin MSComCtl2.DTPicker DTPickerPropertyDateAssigned 
            Height          =   315
            Left            =   2400
            TabIndex        =   53
            Top             =   540
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   105578497
            CurrentDate     =   38300
         End
         Begin VB.Label Label38 
            Caption         =   "Property Code"
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
            Left            =   240
            TabIndex        =   59
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label37 
            Caption         =   "Date Assigned"
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
            Left            =   240
            TabIndex        =   58
            Top             =   570
            Width           =   1335
         End
         Begin VB.Label Label36 
            Caption         =   "Amount Due"
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
            Left            =   240
            TabIndex        =   57
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label35 
            Caption         =   "Start Date"
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
            Left            =   240
            TabIndex        =   56
            Top             =   1830
            Width           =   1575
         End
         Begin VB.Label Label34 
            Caption         =   "Other Details"
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
            Left            =   240
            TabIndex        =   55
            Top             =   2220
            Width           =   1575
         End
         Begin VB.Label Label33 
            Caption         =   "Transaction No"
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
            Left            =   240
            TabIndex        =   54
            Top             =   2550
            Width           =   1335
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Possible Specifications"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74880
         TabIndex        =   43
         Top             =   420
         Width           =   4815
         Begin MSComctlLib.ListView ListALLProperties 
            Height          =   3375
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   5953
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
      Begin VB.Frame Frame12 
         Caption         =   "Actual Specifications"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -67080
         TabIndex        =   41
         Top             =   420
         Width           =   5295
         Begin MSComctlLib.ListView ListActualProperties 
            Height          =   3375
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   5953
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   69
         Top             =   420
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   6588
         View            =   3
         MultiSelect     =   -1  'True
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   70
         Top             =   420
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   6588
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
      Begin MSComctlLib.ListView ListALLLandLords 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   73
         Top             =   360
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   6588
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
         Height          =   3855
         Left            =   -74880
         TabIndex        =   74
         Top             =   360
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   6800
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
         Height          =   2055
         Left            =   -74880
         TabIndex        =   96
         Top             =   1920
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   3625
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   105
         Top             =   360
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   5741
         View            =   3
         MultiSelect     =   -1  'True
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Increment Percent"
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
         Left            =   -72120
         TabIndex        =   104
         Top             =   990
         Width           =   1110
      End
      Begin VB.Label Label26 
         Caption         =   "Payment Due"
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
         Left            =   -74760
         TabIndex        =   103
         Top             =   975
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "Installment"
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
         Left            =   -74760
         TabIndex        =   102
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label txtSerialNo 
         Caption         =   "Serial No"
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
         Left            =   -72120
         TabIndex        =   101
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Current Period"
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
         Left            =   -68760
         TabIndex        =   100
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label42 
         Caption         =   "Contract Year"
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
         Left            =   -68760
         TabIndex        =   99
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "Pay Due Date"
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
         Left            =   -66120
         TabIndex        =   98
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Invoice No"
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
         Left            =   -66120
         TabIndex        =   97
         Top             =   990
         Width           =   675
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
   Begin VB.Menu mnuFaceActions 
      Caption         =   "Face Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteFace 
         Caption         =   "Delete Face?"
      End
   End
   Begin VB.Menu mnuMastActions 
      Caption         =   "Mast Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteMast 
         Caption         =   "Delete Mast?"
      End
   End
   Begin VB.Menu mnuInstalmentActions 
      Caption         =   "Instalment Actions"
      Begin VB.Menu mnuDeleteInstalment 
         Caption         =   "Delete Instalment?"
      End
   End
   Begin VB.Menu mnuContractActions 
      Caption         =   "Contract Actions"
      Begin VB.Menu mnuEditContract 
         Caption         =   "Edit Contract"
      End
      Begin VB.Menu mnuDeleteContract 
         Caption         =   "Delete Contract"
      End
   End
End
Attribute VB_Name = "frmODASMSiteRenewal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSITE As New clsODASSite, MyCommonData As clsCommonData
Dim rsPROPERTY As New clsODASProperties
Public rsLANDLORD As clsODASLandLord
Dim rsALLOCATION As clsODASAllocation

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

Private Sub loadCouncil()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPCouncil WHERE CouncilCode = '" & .cboCouncil.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.RecordCount = 0 Then Exit Sub
        
        .txtCouncil.Text = rsFindRecord!Council
        .cboCouncil.Text = rsFindRecord!CouncilCode
       
    End With
Exit Sub
err:
ErrorMessage
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
         If rsFindRecord.EOF Or rsFindRecord.BOF Then Exit Sub
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

Private Sub cmdcancelContract_Click()
Me.Toolbar1.Enabled = True
rsSITE.loadRECORD
rsLANDLORD.loadLANDLORD
rsSITE.loadMAST
rsSITE.loadSITE

loadPAYMENTMODE
loadCouncil

showALLMASTS
showALLSITES
showALLContracts
showALLINSTALLMENTSDUE

End Sub

Private Sub cmdNewContract_Click()
clearContractDetails
GenerateContractNo
End Sub

Private Sub clearContractDetails()
Me.txtCommencementDate.Text = ""
Me.txtLeaseDuration.Text = ""
Me.txtCommencementDate.Text = ""
Me.txtExpiryDate.Text = ""
Me.cboPaymentMode.Text = ""
Me.txtpaymentMode.Text = ""
Me.txtIncrementFrequency.Text = 1
Me.txtPaymentInterval.Text = 1
Me.txtIncrementStartYear.Text = 2
Me.txtAnnualRentIncrement.Text = 0
Me.txtAnnualRentIncrementType.Text = "N"
Me.txtContractNo.Text = ""
Me.Toolbar1.Enabled = False
End Sub


Private Sub cmdPropertyAdd_Click()
With frmODASMSiteRegistration
        If .txtPlotNo.Text <= "" Then Exit Sub
        If .txtMastNo.Text <= "" Then Exit Sub
        If .txtSiteNo.Text <= "" Then Exit Sub
        
        If Trim(.cmdPropertyADD.Caption) = ">>" Then
                .cmdPropertyADD.Caption = "Save"
                .cmdPropertyDelete.Enabled = False
                rsPROPERTY.clearRECORD
                rsPROPERTY.enableRECORD
                showALLPROPERTIES1
                showACTUALPROPERTIES1

                NewRecord = True
        ElseIf .cmdPropertyADD.Caption = "Save" Then
                j = .ListALLProperties.ListItems.Count: k = 0
                    
                If j = 0 Then Exit Sub
                
                For i = 1 To j
                    If .ListALLProperties.ListItems(i).Checked = True Then
                        k = k + 1
                    End If
                Next i
                        
                If rsPROPERTY.ValidRecord = True Then

                        rsPROPERTY.saveRecord
                        disableALLRECORD
                        showALLPROPERTIES1
                        showACTUALPROPERTIES1
                        disableALLRECORD
                        NewRecord = False
                        .cmdPropertyADD.Caption = ">>"
                        disableALLRECORD
                        .cmdPropertyDelete.Enabled = True
                        .cmdPropertyDelete.Caption = "<<"
                        bsearchRECORD = False
                        beditRECORD = False
                        baddRECORD = False
                        RequireSite = False
                        .Frame2.Enabled = True
                        .SSTab1.Tab = 3
                End If
        End If

End With

End Sub

Private Sub cmdPropertyDelete_Click()
On Error GoTo err
    With frmODASMSiteRegistration
        If .cmdPropertyDelete.Caption = "<<" Then
                
                bDeleteRECORD = True
                .cmdPropertyADD.Enabled = False
                .cmdPropertyDelete.Caption = "Save"
                
        ElseIf .cmdPropertyDelete.Caption = "Save" Then
                 rsPROPERTY.deleteRECORD
                .cmdPropertyADD.Enabled = True
                .cmdPropertyDelete.Caption = "<<"
                .cmdPropertyADD.Caption = "Add"
                showALLPROPERTIES1
                showACTUALPROPERTIES1

        End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdPropertyPrint_Click()

End Sub

Private Sub cmdPropertyRefresh_Click()
        Me.cmdPropertyADD.Enabled = True
        Me.cmdPropertyDelete.Enabled = True
        Me.Toolbar1.Visible = True
End Sub

Private Sub cmdFACESearch_Click()
On Error GoTo err
        With frmODASMSite
                bsearchRECORD = True
                .cmdFaceEDIT.Enabled = True
                .cmdFaceEDIT.Caption = "Edit"
                .SSTab1.Tab = 2
                .Toolbar1.Visible = False
        End With
Exit Sub
err:
    ErrorMessage
End Sub



Private Sub cmdSave_Click()
With Me
        If .cmdSave.Caption = "Change" Then
            .ListALLInstallments.Enabled = True
            .cmdSave.Caption = "Save"
            
            .txtPaymentDue.Locked = False
            .txtInstallmentPercent.Locked = False
            
            
        ElseIf .cmdSave.Caption = "Save" Then
            .cmdSave.Caption = "Change"
            
            If validINSTALLMENT Then
                    saveINSTALLMENT
                    showALLINSTALLMENTSDUE
            End If
        End If
        
End With
End Sub
Private Function validINSTALLMENT()
On Error GoTo err
    With frmODASMSiteRegistration
            validINSTALLMENT = False
            
            If .txtTransactionNo.Text <= 0 Then
                MsgBox "The Transaction Number is Required ..................."
                .txtTransactionNo.SetFocus
                
            ElseIf .txtInstallment.Text <= 0 Then
                MsgBox "The Installment No cannot be Blank .................."
                .txtInstallment.SetFocus
                
            ElseIf .txtPaymentDue.Text < 0 Then
                MsgBox "The Payment Due Must be > Zero................."
                .txtPaymentDue.SetFocus
            
            ElseIf IsDate(.txtPaymentDueDate.Text) <> True Then
                MsgBox "The Payment Due Date Captured is Invalid ........."
                .txtPaymentDueDate.SetFocus
            
            ElseIf .txtCurrentPeriod.Text <= "" Then
                MsgBox "The Current Period Cannot be Left Blank .........."
                .txtCurrentPeriod.SetFocus
            
            ElseIf .txtContractYear.Text <= "" Then
                MsgBox "The Current Year is invalid ....................."
                .txtContractYear.SetFocus
            Else
                    validINSTALLMENT = True
            End If

    End With
Exit Function
err:
    ErrorMessage
End Function


Private Sub saveINSTALLMENT()
On Error GoTo err
    With frmODASMSiteRegistration
    
            Set rsSAVE = New ADODB.Recordset
            rsSAVE.Open "SELECT * FROM ODASMInstallment WHERE InstallmentNo = '" & .txtTransactionNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.EOF And rsSAVE.BOF Then
            Else
                    rsSAVE!ContractYear = .txtContractYear
                    rsSAVE!TotalRent = .txtPaymentDue
                    rsSAVE!PaymentDueDate = .txtPaymentDueDate
                    rsSAVE!CurrentPeriod = CurrentPeriod
                    rsSAVE!InstallmentPercent = .txtInstallmentPercent
                    rsSAVE!PaymentDue = .txtPaymentDue
                    rsSAVE!Balance = .txtPaymentDue
                    
                    If .chkPaymentFlag.Value = 1 Then
                        rsSAVE!PaymentFlag = "Y"
                    Else: rsSAVE!PaymentFlag = "N"
                    End If
                    
                    rsSAVE.Update
                    Set rsSAVE = Nothing
                    
            End If
            

    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdSaveContract_Click()
 If rsSITE.validFACE Then
    rsALLOCATION.SaveContractDetails
    rsALLOCATION.GenerateInstallmentPayable
End If
End Sub

Private Sub DTPickerCommencementDate_CloseUp()
On Error GoTo err
With Me
        .txtCommencementDate.Text = .DTPickerCommencementDate.Value
        .txtExpiryDate.Text = DateAdd("yyyy", CDbl(.txtLeaseDuration), .txtCommencementDate.Text)
        .txtExpiryDate.Text = DateAdd("D", -1, .txtExpiryDate)
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub DTPickerPropertyCommencementDate_CloseUp()
On Error GoTo err
    With frmODASMSiteRegistration
        .txtPropertyCommencementDate.Text = .DTPickerPropertyCommencementDate
    End With

Exit Sub
err:
    ErrorMessage
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

Private Sub Form_Activate()
        disableALLRECORD
        showALLPLOTS
        showALLPROPERTIES1
        showACTUALPROPERTIES1

End Sub

Private Sub Form_Initialize()
        Set rsSITE = New clsODASSite
        Set rsPROPERTY = New clsODASProperties
        Set rsLANDLORD = New clsODASLandLord
        Set rsALLOCATION = New clsODASAllocation
End Sub

Private Sub Form_Load()
        rsSITE.loadDEFAULTS
        showALLLandlords
        Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Terminate()
        Set rsSITE = Nothing
        Set rsPROPERTY = Nothing
        Set rsLANDLORD = Nothing
        Set rsALLOCATION = Nothing
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

Private Sub ListActualProperties_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListActualProperties.SortKey = ColumnHeader.Index - 1
    ListActualProperties.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListActualProperties_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListActualProperties.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListActualProperties.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListActualProperties.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMSiteRegistration.txtPropertyCode.Text = Item.Text
            rsPROPERTY.loadRECORD
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListALLInstallments_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If ListALLInstallments.View = lvwList Or Me.ListALLInstallments.ListItems.Count = 0 Then Exit Sub
    If Button = 2 Then
        PopupMenu mnuInstalmentActions
    End If
End Sub

Private Sub ListALLLandLords_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListALLLandLords.SortKey = ColumnHeader.Index - 1
    ListALLLandLords.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListALLInstallments_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListALLInstallments.SortKey = ColumnHeader.Index - 1
    ListALLInstallments.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListALLInstallments_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListALLInstallments.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListALLInstallments.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListALLInstallments.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMSiteRegistration.txtTransactionNo.Text = Item.Text
            loadINSTALLMENT
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadINSTALLMENT()
On Error GoTo err
    With frmODASMSiteRegistration
               
            Set rsCONTROL = New ADODB.Recordset
            rsCONTROL.Open "SELECT * FROM ODASMInstallment  WHERE InstallmentNo = '" & .txtTransactionNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            If IsNull(rsCONTROL!PaymentDue) = True Then
                .txtPaymentDue.Text = 0
            Else: .txtPaymentDue.Text = FormatNumber(rsCONTROL!PaymentDue, 2)
            End If
            
            If IsNull(rsCONTROL!InstallmentPercent) = True Then
                .txtInstallmentPercent.Text = 0
            Else: .txtInstallmentPercent.Text = FormatNumber(rsCONTROL!InstallmentPercent, 2)
            End If
            
            
            .txtInstallment.Text = rsCONTROL!Installment & ""
            .txtInvoiceNo.Text = rsCONTROL!InvoiceNo & ""
            
            If IsDate(rsCONTROL!PaymentDueDate) Then .txtPaymentDueDate.Text = rsCONTROL!PaymentDueDate
            
            .txtCurrentPeriod.Text = rsCONTROL!CurrentPeriod & ""
            .txtContractYear.Text = rsCONTROL!ContractYear & ""
            
            
            Set rsCONTROL = Nothing

    
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListALLLandLords_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListALLLandLords.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListALLLandLords.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListALLLandLords.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMSiteRegistration.txtLandLordNo.Text = Item.Text
            rsLANDLORD.loadLANDLORD
             
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
            
            frmODASMSiteRegistration.txtPlotNo.Text = Item.Text
            
            
            rsSITE.loadRECORD
            loadCouncil
            loadPAYMENTMODE
            showALLMASTS
            showALLSITES
            showALLINSTALLMENTSDUE
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListALLProperties_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListALLProperties.SortKey = ColumnHeader.Index - 1
    ListALLProperties.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListALLProperties_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListALLProperties.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            frmODASMSiteRegistration.txtPropertyCode.Text = Item.Text
            
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

            rsSITE.loadMAST frmODASMSiteRegistration.txtMastNo.Text
            showALLSITES
            showALLINSTALLMENTSDUE
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If ListView2.View = lvwList Or Me.ListView2.ListItems.Count = 0 Then Exit Sub
    If Button = 2 Then
        PopupMenu mnuMastActions
    End If
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
            frmODASMSiteRegistration.txtMastNo.Text = Item.SubItems(1)

            rsSITE.loadSITE frmODASMSiteRegistration.txtSiteNo.Text
            rsSITE.loadMAST frmODASMSiteRegistration.txtMastNo.Text
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Option1_Click()

End Sub

Private Sub ListView3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If ListView1.View = lvwList Or Me.ListView1.ListItems.Count = 0 Then Exit Sub
    If Button = 2 Then
        PopupMenu mnuFaceActions
    End If
End Sub

Private Sub ListView4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If ListView4.View = lvwList Or Me.ListView4.ListItems.Count = 0 Then Exit Sub
    If Button = 2 Then
        PopupMenu mnuInstalmentActions
    End If

End Sub

Private Sub mnuDeleteFace_Click()
DeleteFace Me.ListView3.SelectedItem.Text
'DeleteMast (MastNo)
End Sub

Private Sub DeleteFace(SiteNo As String)
On Error GoTo errMSG
        If MsgBox("Are you sure you want to delete this face?", vbYesNo) = vbYes Then
            strSQL = "DELETE FROM ODASPPlotSite WHERE PlotNo='" & Me.txtPlotNo.Text & "' AND SiteNo LIKE '" & SiteNo & "'"
            Set rsDeleteRecord = cnCOMMON.Execute(strSQL)
            rsSITE.updateALLSites
            showALLSITES
            showALLMASTS
            
        End If
Exit Sub
errMSG:
        ErrorMessage

End Sub

Private Sub DeleteInstalment(InstalmentNo As String)
On Error GoTo errMSG
        If MsgBox("Are you sure you want to delete this Instalment?", vbYesNo) = vbYes Then
            strSQL = "DELETE FROM ODASMInstallment WHERE InstallmentNo LIKE '" & InstalmentNo & "' "
            Set rsDeleteRecord = cnCOMMON.Execute(strSQL)
            rsSITE.updateALLSites
            showALLINSTALLMENTSDUE
            
        End If
Exit Sub
errMSG:
        ErrorMessage

End Sub

Private Sub DeleteMast(MastNo As String)
On Error GoTo errMSG
        If MsgBox("Are you sure you want to delete this Mast?", vbYesNo) = vbYes Then
            strSQL = "DELETE FROM ODASPPlotSite WHERE PlotNo LIKE '" & Me.txtPlotNo.Text & "' AND MastNo LIKE '" & MastNo & "';"
            strSQL = strSQL & "DELETE FROM ODASPPLotMast WHERE PlotNo='" & Me.txtPlotNo.Text & "' AND MastNo LIKE '" & MastNo & "' "
            Debug.Print strSQL
            Set rsDeleteRecord = cnCOMMON.Execute(strSQL)
            
            rsSITE.updateANNUALRate
            rsSITE.updateANNUALRent
            rsSITE.updateALLSites
            showALLSITES
            showALLMASTS
            
        End If
Exit Sub
errMSG:
        ErrorMessage

End Sub


Private Sub mnuDeleteInstalment_Click()
DeleteInstalment Me.ListALLInstallments.SelectedItem.Text
End Sub

Private Sub mnuDeleteMast_Click()
DeleteMast Me.ListView2.SelectedItem.Text

End Sub

Private Sub mnuEditContract_Click()
Me.SSTab1.Tab = 1
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
    If NewRecord <> True Then Exit Sub
    If .optYES.Value = True Then Exit Sub
    .txtMastAnnualRent.BackColor = &HFFC0C0
    .txtMastAnnualRent.Locked = False
End With
End Sub
Private Sub optNo_Click()
With Me
    If NewRecord <> True And beditRECORD <> True Then Exit Sub
    .txtMastAnnualRent.Locked = False
    .txtMastAnnualRent.BackColor = &HFFC0C0
    .txtPlotNo.Text = Empty
End With
End Sub
Private Sub optYES_Click()
With Me
    .txtMastAnnualRent.Locked = True
    .txtMastAnnualRent.Text = 0
    .txtMastAnnualRent.BackColor = &HFFFFC0
End With
End Sub

Private Sub clearRECORD()
On Error GoTo err
    With Me
            .txtAnnualRentIncrement.Text = 0
            .txtNoofMasts.Text = 0
            .txtNoofSites.Text = 0
            .txtIncrementStartYear = 2
            .txtIncrementFrequency = 1
            .txtPaymentInterval = 1
            .txtAnnualRentIncrementType.Text = "N"
            .txtRentVariationType.Text = "N"
            .optEAGLE.Value = True
            .optYES.Value = True
            .DTPickerAcquisitionDate.Value = Date
            .DTPickerCommencementDate.Value = Date
            .DTPickerPropertyCommencementDate.Value = Date
            .DTPickerPropertyDateAssigned.Value = Date
            .txtLeaseDuration.Text = 0
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        
        With frmODASMSiteRegistration
        
        Select Case Button.Key
        Case "N"
            Select Case Button.Caption
                Case "New &Record "
                    
                    If editRECORD Then Exit Sub
                    rsSITE.enablePLOT: rsSITE.enableOptions: rsSITE.enableSITE
                    NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                    clearRECORD
                Case "&Save Record "
                    If rsSITE.validFACE Then
                            
                            bAllowProcess = True
                            rsLANDLORD.updateRECORDExternally
                            bAllowProcess = False

                            rsSITE.updateRECORD
                            rsSITE.saveMAST
                            rsSITE.saveSITE
                        
                            rsSITE.updateANNUALRate
                            rsSITE.updateANNUALRent
                            rsSITE.updateALLSites
                            rsALLOCATION.SaveContractDetails
                            rsALLOCATION.GenerateInstallmentPayable

                            showALLMASTS
                            showALLPLOTS
                            showALLSITES
                            showALLINSTALLMENTSDUE
                            showALLLandlords
                            showALLPROPERTIES1
                            showACTUALPROPERTIES1

                            disableALLRECORD
                            NewRecord = False
                            
                            Button.Caption = "NEXT &MAST": Button.Image = 2
                            .Toolbar1.Buttons(3).Caption = "NEXT &FACE": .Toolbar1.Buttons(3).Image = 2
                            .Toolbar1.Buttons(4).Caption = "FINISH"
                    End If

                 Case "NEXT &MAST"
                        rsSITE.clearSITE
                        rsSITE.enableSITE
                        rsSITE.clearMAST
                        .optCLIENT.Enabled = True
                        .optEAGLE.Enabled = True
                        .optCLIENT.Value = False
                        If .txtAnnualRent > 0 Then
                            .optEAGLE.Value = True
                        End If
                       
                        .SSTab1 = 0
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
                                editRECORD = True
                            End If
                Case "NEXT &FACE"
                        rsSITE.clearSITE
                        rsSITE.enableSITE
                        NewRecord = True: .Toolbar1.Buttons(2).Image = 4: .Toolbar1.Buttons(2).Caption = "&Save Record ": bSaveRECORD = False
                
                Case "Save &Changes "
                    If rsSITE.validFACE Then
                            rsSITE.saveRecord
                            rsSITE.saveMAST
                            rsSITE.saveSITE
                            
                            rsSITE.updateANNUALRate
                            rsSITE.updateANNUALRent
                            rsSITE.updateALLSites
                            
                            
                            bAllowProcess = True
                            rsLANDLORD.updateRECORDExternally
                            bAllowProcess = False
                            
                            rsALLOCATION.GenerateInstallmentPayable

                            showALLMASTS
                            showALLPLOTS
                            showALLSITES
                            showALLPROPERTIES1
                            showACTUALPROPERTIES1
                            showALLINSTALLMENTSDUE
                            showALLLandlords
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
                            NewRecord = False: editRECORD = fasle
                Case Else
                   
                    Exit Sub
                End Select
        Case "S"
                Select Case Button.Caption
                        Case "&Search/Find "
                        .SSTab1.Tab = 2
                        CurrentRecord = InputBox("Enter the Plot number to search ...")
                        .Toolbar1.Buttons(2).Caption = "New &Record "
                        .Toolbar1.Buttons(2).Image = 2
                        .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                        .Toolbar1.Buttons(3).Image = 5

                        If Len(CurrentRecord) = 0 Then Exit Sub
                        .txtPlotNo.Text = CurrentRecord
                        rsSITE.loadRECORD
                        rsLANDLORD.loadLANDLORD
                        rsSITE.loadMAST
                        rsSITE.loadSITE
                        
                        loadPAYMENTMODE
                        loadCouncil
                        
                        showALLMASTS
                        showALLSITES
                        showALLContracts
                        showALLINSTALLMENTSDUE
                        Case "FINISH"
                            Unload Me

                Case "Cancel"
                    cancelCMD
                Case "Delete"
                    'cmdDelete_Click
                    Button.Caption = "&Search/Find "

                Case Else
                   
                    Exit Sub
                End Select
        Case "D"
                If Trim(Me.txtPlotNo.Text) = "" Then
                        MsgBox "Select a site to delete", vbExclamation
                        Exit Sub
                End If
                If MsgBox("Are you sure you want to delete this record", vbYesNo + vbExclamation) = vbYes Then
                        cmdDelete_Click
                End If
        Case "R"
                If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 2
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 5
                    NewRecord = False: editRECORD = False: rsSITE.clearRECORD
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

Exit Sub
err:
    ErrorMessage

End Sub


Private Sub cmdDelete_Click()
On Error GoTo errMSG
strSQL = "DELETE FROM ODASPPlot WHERE PlotNo LIKE '" & Me.txtPlotNo.Text & "';"
strSQL = strSQL & "DELETE FROM ODASPPlotMast WHERE PlotNo LIKE '" & Me.txtPlotNo.Text & "';"
strSQL = strSQL & "DELETE FROM ODASPPlotSite WHERE PlotNo LIKE '" & Me.txtPlotNo.Text & "';"
strSQL = strSQL & "DELETE FROM ODASMInstallment WHERE ContractNo LIKE '" & Me.txtPlotNo.Text & "';"

Set rsDeleteRecord = cnCOMMON.Execute(strSQL)

Toolbar1.Buttons(2).Caption = "New &Record "
Toolbar1.Buttons(2).Image = 2
Toolbar1.Buttons(3).Caption = "&Edit/Change "
Toolbar1.Buttons(3).Image = 5
clearALLRECORD
Exit Sub
errMSG:
        ErrorMessage
End Sub

Private Sub loadExistingRECORD()
      With frmODASMSiteRegistration
                    INQUIRY = InputBox("Enter the Plot Number to search and display...", "Search Value request")
                    If Len(INQUIRY) = 0 Then Exit Sub
                    strSQL = "select * from ODASPPlot Where ODASPPlot.PlotNo = '" & INQUIRY & "' ;"
                    Set rsFind = New ADODB.Recordset
                    rsFind.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    If rsFind.RecordCount = 0 Then
                        MsgBox "Requested Plot could not be found. Either it was deleted or you entered a wrong Number", vbCritical + vbOKOnly, "Records Search Engine"
                    Else
                           .txtPlotNo.Text = rsFind!PlotNo
                           .txtAcquiredBy.Text = rsFind!CreatedBy
                           .txtStatus.Text = rsFind!Status
                           .txtTownCode.Text = rsFind!TownCode
                           .txtLeaseDuration.Text = rsFind!LeaseDuration
                           .txtLRNo.Text = rsFind!LRNo
                           .txtPhysicalAddress.Text = rsFind!PhysicalLocation
                           .txtCommencementDate.Text = rsFind!CommencementDate
                    
                    Set rsFind = Nothing
                    strSQL = Empty
                            showALLMASTS
                            showALLSITES
                            showALLINSTALLMENTSDUE
                            .Toolbar1.Buttons(2).Caption = "NEXT &MAST"
                            .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                            Button.Caption = "FINISH"
                    End If
                    
           End With

End Sub

Private Sub updatePLOTSITE()
On Error GoTo err
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
                  
                  If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
        
                  With rsSAVE
                  
                    .MoveFirst
                          Do While Not .EOF
                                !ApprovedBy = CurrentUserName
                                !DateApproved = Date
                                !Approved = "Y"
                                
                                !AuthorizedBy = CurrentUserName
                                !DateAuthorized = Date
                                        
                                !Authorized = "Y"
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
                           
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        
                        !Approved = "Y"
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date
                        
                        !Authorized = "Y"
                        !Status = "SITE-AVAILABLE"
                        !RentDueDate = CDbl(!CommencementDate)
                        !RentDue = CDbl(!AnnualRent)
                                
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

Private Sub txtLandLordName_Change()
bsearchRECORD = True
Me.SSTab1.Tab = 5
showALLLandlords
bsearchRECORD = False
End Sub


Private Sub txtLeaseDuration_LostFocus()
UpDownLeasePeriod_Change
End Sub

Private Sub txtMastAnnualRent_LostFocus()
        If Me.txtMastAnnualRent.Text = "" Then Exit Sub
        Me.txtMastAnnualRent.Text = FormatNumber(Me.txtMastAnnualRent.Text, 2)
End Sub

Private Sub txtPhysicalAddress_Change()
bsearchRECORD = True
Me.SSTab1.Tab = 2
showALLPLOTS
'showALLPROPERTIES1
'showACTUALPROPERTIES1

bsearchRECORD = False
End Sub

Private Sub txtPhysicalAddress_LostFocus()
With Me
    .txtMastDetails.Text = UCase(.txtPhysicalAddress.Text) & ""
    .txtPhysicalAddress.Text = UCase(.txtPhysicalAddress.Text) & ""
    showALLPLOTS
End With
End Sub


Private Sub txtPropertyAmountDue_lOSTfOCUS()
        Me.txtPropertyAmountDue.Text = FormatNumber(Me.txtPropertyAmountDue.Text, 2)
End Sub

Private Sub txtSiteNo_LostFocus()
    showALLMASTS
End Sub
Private Sub UpDownLeasePeriod_Change()
On Error GoTo err
        With frmODASMSiteRegistration
            .txtLeaseDuration.Text = .UpDownLeasePeriod.Value
            .txtExpiryDate.Text = DateAdd("YYYY", .UpDownLeasePeriod.Value, Format(.txtCommencementDate.Text, "MMMM dd,yyyy"))
        .txtExpiryDate.Text = DateAdd("yyyy", CDbl(.txtLeaseDuration), .txtCommencementDate.Text)
        .txtExpiryDate.Text = DateAdd("D", -1, .txtExpiryDate)
        
        End With

Exit Sub

err:
    ErrorMessage
End Sub
