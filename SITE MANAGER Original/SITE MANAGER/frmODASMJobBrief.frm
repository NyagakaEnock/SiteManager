VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMJobBrief 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internal Job Brief"
   ClientHeight    =   7995
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmODASMJobBrief.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Information Attached to this Brief"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   61
      Top             =   6120
      Width           =   5175
      Begin MSComctlLib.ListView ListView1 
         Height          =   1335
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
   Begin VB.TextBox txtQuotationNo 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   4320
      TabIndex        =   47
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtJobBriefDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2520
      TabIndex        =   39
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame9 
      Caption         =   "Job Brief Details"
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
      Height          =   1695
      Left            =   5400
      TabIndex        =   32
      Top             =   6120
      Width           =   5895
      Begin MSComctlLib.ListView ListView5 
         Height          =   1335
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
   Begin VB.Frame FrameSites 
      Caption         =   "Available Sites"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   30
      Top             =   4080
      Width           =   5175
      Begin VB.CheckBox chkAcquireSite 
         Caption         =   "Acquire Site?"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox txtSiteLocation 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   53
         Top             =   1680
         Width           =   3615
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   1455
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
   Begin VB.Frame Frame6 
      Caption         =   "Project Details"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   5175
      Begin VB.TextBox txtReceivedBy 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   3720
         TabIndex        =   58
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtBriefBy 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   57
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtExpectedDOC 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtdeadlineDate 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   3720
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAdditionalInformation 
         BackColor       =   &H00FFC0C0&
         Height          =   555
         Left            =   1080
         MaxLength       =   249
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   960
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtpDeadlineDate 
         Height          =   315
         Left            =   4800
         TabIndex        =   23
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57475073
         CurrentDate     =   38288
      End
      Begin MSComCtl2.DTPicker DTPickerExpectedDOC 
         Height          =   315
         Left            =   2280
         TabIndex        =   43
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57475073
         CurrentDate     =   38288
      End
      Begin VB.Label Label14 
         Caption         =   "Received By"
         Height          =   255
         Left            =   2640
         TabIndex        =   60
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Brief By"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Exp DOC"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Additional Information"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Deadline Date"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Varies Sizes Of The Media"
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
      Height          =   1575
      Left            =   5400
      TabIndex        =   20
      Top             =   2520
      Width           =   5895
      Begin MSComctlLib.ListView ListView3 
         Height          =   1215
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2143
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
   Begin MSComCtl2.DTPicker dtpBriefDate 
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57475073
      CurrentDate     =   38283
   End
   Begin VB.TextBox txtJobBriefNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Item(s) Quoted"
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
      Height          =   2055
      Left            =   5400
      TabIndex        =   4
      Top             =   4080
      Width           =   5895
      Begin VB.TextBox txtPhysicalLocation 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   65
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtAttachmentCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2760
         TabIndex        =   64
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtMediaSize 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   960
         TabIndex        =   49
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtJobBriefItemNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboColorCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4320
         TabIndex        =   37
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtSiteNo 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtSiteName 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   960
         TabIndex        =   34
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtMediaCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2760
         TabIndex        =   29
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox chkBorder 
         Caption         =   "Bordered"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1740
         Width           =   975
      End
      Begin VB.CheckBox chkIlluminate 
         Caption         =   "Install Lights?"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   1710
         Width           =   1335
      End
      Begin VB.TextBox txtItemQuantity 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4920
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox cboSidingCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4320
         TabIndex        =   15
         Top             =   1680
         Width           =   1455
      End
      Begin MSComCtl2.UpDown UpDownQuantity 
         Height          =   315
         Left            =   5520
         TabIndex        =   52
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Max             =   99
         Enabled         =   -1  'True
      End
      Begin VB.Label Label15 
         Caption         =   "Location"
         Height          =   255
         Left            =   3120
         TabIndex        =   66
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Attachments"
         Height          =   255
         Left            =   3120
         TabIndex        =   63
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Site Details"
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Size"
         Height          =   255
         Left            =   1560
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Color "
         Height          =   255
         Left            =   4860
         TabIndex        =   38
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Site "
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "Qty"
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Siding"
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Item No"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Media"
         Height          =   255
         Left            =   3300
         TabIndex        =   12
         Top             =   240
         Width           =   495
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
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   55
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         MaxLength       =   120
         TabIndex        =   41
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtCustomerID 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtContactName 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Product"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   2160
         TabIndex        =   56
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Description"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Contact"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   " Name"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Active Media Code"
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
      Left            =   5400
      TabIndex        =   1
      Top             =   600
      Width           =   5895
      Begin MSComctlLib.ListView ListView2 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2778
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
            Picture         =   "frmODASMJobBrief.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMJobBrief.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMJobBrief.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMJobBrief.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMJobBrief.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMJobBrief.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMJobBrief.frx":236E
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
      Width           =   11565
      _ExtentX        =   20399
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
            Caption         =   "&Help"
            Key             =   "F"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
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
   Begin VB.Label Label17 
      Caption         =   "Q/ No"
      Height          =   255
      Left            =   3840
      TabIndex        =   48
      Top             =   750
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   " Date"
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
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   750
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   " Brief No"
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   750
      Width           =   1215
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
      Begin VB.Menu mnuRegisteredClients 
         Caption         =   "Registered Clients"
      End
      Begin VB.Menu mnuKHJGGFDHJ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowQuotations 
         Caption         =   "Show Quotations"
      End
      Begin VB.Menu mnuExtraInfo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtraInformation 
         Caption         =   "Extra Inform"
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
Attribute VB_Name = "frmODASMJobBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCOMPANY As clsODASJobBrief

Private Sub cboSiding_GotFocus()
        SelectSidingGotFocus
End Sub

Private Sub cboSiding_KeyPress(KeyAscii As Integer)
        KeyPress = 0
End Sub

Private Sub cboSiding_LostFocus()
        selectSidingLostFocus
End Sub

Private Sub cboColorCode_GotFocus()
        SelectColorGotFocus
End Sub

Private Sub cboColorCode_KeyPress(KeyAscii As Integer)
        KeyPress = 0
End Sub

Private Sub cboColorCode_LostFocus()
        selectColorLostFocus
End Sub

Private Sub cboDiscountCode_GotFocus()
        SelectDiscountGotFocus
End Sub

Private Sub cboDiscountCode_KeyPress(KeyAscii As Integer)
        KeyPress = 0
End Sub

Private Sub cboDiscountCode_LostFocus()
        selectDiscountLostFocus
End Sub

Private Sub cboSidingCode_Gotfocus()
    SelectSidingGotFocus
End Sub

Private Sub cboSidingCode_KeyPress(KeyAscii As Integer)
    KeyPress = 0
End Sub

Private Sub cboSidingCode_LostFocus()
    selectSidingLostFocus
End Sub

Private Sub chkAcquireSite_Click()
On Error GoTo err
        With frmODASMJobBrief
            
            If .txtMediaCode.Text = Empty Then Exit Sub
            
            If .chkAcquireSite.Value = 1 Then
                .txtSiteLocation.Locked = False
                .txtSiteLocation.BackColor = &HFFC0C0
            Else
                .txtSiteLocation.Locked = True
                .txtSiteLocation.BackColor = &HFFFFC0
            End If
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub dtpBriefDate_Change()
'On Error GoTo err
        With frmODASMJobBrief
            .txtJobBriefDate.Text = .dtpBriefDate.Value
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub


Private Sub dtpDeadlineDate_Change()
'On Error GoTo err
        With frmODASMJobBrief
            .txtDeadlineDate.Text = .dtpDeadlineDate.Value
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub DTPickerExpectedDOC_Change()
'On Error GoTo err
        With frmODASMJobBrief
            .txtExpectedDOC.Text = .DTPickerExpectedDOC.Value
        End With
        
Exit Sub
err:
    ErrorMessage

End Sub
Private Sub DTPickerCommencementDate_Change()
'On Error GoTo err
        With frmODASMJobBrief
            .txtCommencementDate.Text = .DTPickerCommencementDate.Value
        End With
        
Exit Sub
err:
    ErrorMessage

End Sub
Private Sub DTPickerExpiryDate_Change()
'On Error GoTo err
        With frmODASMJobBrief
            .txtExpiryDate.Text = .DTPickerExpiryDate.Value
        End With
        
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub Form_Activate()
    Set rsCOMPANY = New clsODASJobBrief

        rsCOMPANY.loadDEFAULTS
        rsCOMPANY.loadCONTACT
        disableLISTVIEW
        Set rsCOMPANY = Nothing
        showALLMEDIA2
        showJOBBRIEFITEMS
        showALLATTACHMENT
        disableALLRECORD

    Set rsCOMPANY = Nothing
End Sub
Private Sub enableLISTVIEW()
On Error GoTo err
        
        With frmODASMJobBrief
            .ListView1.Enabled = True
            .ListView2.Enabled = True
            .ListView3.Enabled = True
            .ListView4.Enabled = True
            .ListView5.Enabled = True
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub disableLISTVIEW()
On Error GoTo err
        
        With frmODASMJobBrief
            .ListView1.Enabled = False
            .ListView2.Enabled = False
            .ListView3.Enabled = False
            .ListView4.Enabled = False
            .ListView5.Enabled = False
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub disableFRAMES()
'On Error GoTo err
        
        With frmODASMJobBrief
            .Frame9.Enabled = False
            .Frame2.Enabled = False
            .Frame6.Enabled = False
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Initialize()
        Set rsCOMPANY = New clsODASJobBrief
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMJobBrief.txtAttachmentCode.Text = Item.Text

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Terminate()
       Set rsCOMPANY = Nothing
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
'On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMJobBrief.txtMediaCode.Text = Item.Text

            Set rsCOMPANY = New clsODASJobBrief
            rsCOMPANY.loadMEDIADETAILS
            Set rsCOMPANY = Nothing
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView3.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView3.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView3.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMJobBrief.txtMediaSize.Text = Item.Text

            Set rsCOMPANY = New clsODASJobBrief
            showUNALLOCATEDSites
            Set rsCOMPANY = Nothing
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView4.SortKey = ColumnHeader.Index - 1
    ListView4.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView4_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView4.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView4.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView4.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMJobBrief.txtSiteNo.Text = Item.Text

            Set rsCOMPANY = New clsODASJobBrief
            rsCOMPANY.loadSiteName
            rsCOMPANY.calculatePRICE
            Set rsCOMPANY = Nothing
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuExtraInformation_Click()
        Load frmODASMInformation
        frmODASMInformation.Show 1, Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo err
Set rsCOMPANY = New clsODASJobBrief
   
 With frmODASMJobBrief
    Select Case Button.Key
    Case "N"
    
                Select Case Button.Caption
                    Case "New &Record "
                        If EditRecord Then Exit Sub
                        disableALLRECORD
                        rsCOMPANY.enableRECORD
                        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                        enableLISTVIEW
                        
                    Case "&Save Record "
                        If NewRecord Then
                                If rsCOMPANY.ValidRecord Then
                                    rsCOMPANY.GenerateJobBriefNo
                                    rsCOMPANY.saveJobBrief
                                    rsCOMPANY.GenerateJobItemNo
                                    rsCOMPANY.saveJobBriefItems
                                    rsCOMPANY.saveJobBriefATTACHMENT
                                    rsCOMPANY.updateMAST
                                    rsCOMPANY.updateSITE
                                    showJOBBRIEFITEMS
                                    showALLATTACHMENT
                                    disableALLRECORD
'                                    RemoveCurrentList3Item
'                                    RemoveCurrentList2Item
                                    Button.Caption = "NE&XT ITEM"
                                    .Toolbar1.Buttons(4).Caption = "NEXT ATTACH"

                                    .Toolbar1.Buttons(3).Caption = "FINISH"
                                
                                End If
                        Else
                                If rsCOMPANY.ValidRecord Then
                                    rsCOMPANY.saveJobBrief
                                    rsCOMPANY.GenerateJobItemNo
                                    rsCOMPANY.saveJobBriefItems
                                    showJOBBRIEFITEMS

                                    Button.Caption = "NE&XT ITEM"
                                    .Frame3.Enabled = False
                                    .Frame2.Enabled = False
                                    .Frame6.Enabled = False
                                    .Toolbar1.Buttons(3).Caption = "FINISH"
                                    .Toolbar1.Buttons(4).Caption = "NEXT ATTACH"

                                End If

                        End If
                        
                     Case "NE&XT ITEM"
                            Button.Caption = "&Save Record ": Button.Image = 4
                            disableFRAMES
                            rsCOMPANY.clearRECORD
                            rsCOMPANY.enableRECORD

                            NewRecord = False
                    Case Else
                            Exit Sub
                End Select
    
      Case "E"
      Case "S"
               Select Case Button.Caption
                    Case "NEXT ATTACH"
                            Button.Caption = "&Save Record ": Button.Image = 4
                            disableFRAMES
                            enableLISTVIEW
                            showALLATTACHMENT
                    
                    Case "&Save Record "
                                If rsCOMPANY.ValidRecord Then
                                    rsCOMPANY.saveJobBriefATTACHMENT
                                    showALLATTACHMENT
                                    disableALLRECORD
                                    .Toolbar1.Buttons(4).Caption = "NEXT ATTACH"
                                    .Toolbar1.Buttons(3).Caption = "FINISH"
                                End If
                End Select
                
        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
        Case "P"
                Load frmODASRJobBrief
                frmODASRJobBrief.Show 1, Me
             
        Case Else
            Exit Sub
        End Select

    End With
Set rsCOMPANY = Nothing

Exit Sub
err:
    ErrorMessage

End Sub

Private Sub txtItemPrice_LostFocus()
'On Error GoTo err
    
    With frmODASMJobBrief
        .txtNetItemPrice.Text = CDbl(.txtItemQuantity.Text) * CDbl(.txtItemPrice.Text)
        .txtNetItemPrice.Text = CCur(.txtNetItemPrice.Text)
    End With

Exit Sub
err:
    ErrorMessage
End Sub


Private Sub UpDownDuration_Change()
'On Error GoTo err
        With frmODASMJobBrief
            .txtDuration.Text = .UpDownDuration.Value
            .txtExpiryDate.Text = DateAdd("M", CDbl(.txtDuration), .txtCommencementDate.Text)
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub UpDown1_Change()

End Sub

Private Sub UpDownQuantity_Change()
'On Error GoTo err
        With frmODASMJobBrief
            .txtItemQuantity.Text = .UpDownQuantity.Value
        End With
Exit Sub

err:
    ErrorMessage
End Sub
