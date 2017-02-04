VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCreateJobBrief 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internal Job Brief"
   ClientHeight    =   8340
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmCreateJobBreif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Caption         =   "Site Details"
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
      Left            =   5760
      TabIndex        =   54
      Top             =   6480
      Width           =   5535
      Begin VB.TextBox txtOtherSite 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   64
         Top             =   960
         Width           =   4335
      End
      Begin VB.CheckBox chkReserve 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reserve"
         Height          =   195
         Left            =   4080
         TabIndex        =   62
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkSiteAvailable 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Available"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   61
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chkSiteAvailable 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Availble"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtSitePhysicalAddress 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   59
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtSideCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4560
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboSiteName 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   56
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label30 
         Caption         =   "Other Site"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Physical Add."
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Site Name"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   1935
      Left            =   5760
      TabIndex        =   39
      Top             =   4560
      Width           =   6135
      Begin MSComCtl2.DTPicker dtpCommenceDate 
         Height          =   375
         Left            =   1200
         TabIndex        =   53
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   38288
      End
      Begin VB.TextBox txtJobBriefComments 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   51
         Top             =   1560
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker dtpExpectedDOC 
         Height          =   375
         Left            =   4200
         TabIndex        =   48
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   38288
      End
      Begin VB.ComboBox cboInfoAttached 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1200
         TabIndex        =   46
         Top             =   1080
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpDeadlineDate 
         Height          =   375
         Left            =   4200
         TabIndex        =   44
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   38288
      End
      Begin VB.TextBox txtLPONo 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4200
         TabIndex        =   43
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtQuotationNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1200
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "Commence"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Expcted D.O.C"
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Info Attached"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Deadline Date"
         Height          =   255
         Left            =   3000
         TabIndex        =   45
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "L.P.O No"
         Height          =   255
         Left            =   3000
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "QuotationNo"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "List Of Items Under Selected Quotation"
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
      TabIndex        =   37
      Top             =   4440
      Width           =   5535
      Begin MSComctlLib.ListView ListView3 
         Height          =   1695
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
      Height          =   255
      Left            =   10080
      TabIndex        =   10
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   19660801
      CurrentDate     =   38283
   End
   Begin VB.TextBox txtJobBriefNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   7080
      TabIndex        =   8
      Top             =   720
      Width           =   1815
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
      Height          =   2535
      Left            =   5760
      TabIndex        =   6
      Top             =   2040
      Width           =   6135
      Begin VB.CheckBox chkBorder 
         Caption         =   "Bordered"
         Height          =   195
         Left            =   1200
         TabIndex        =   36
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox chkIlluminate 
         Caption         =   "Illuminated"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   1920
         Width           =   4935
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   4800
         Max             =   0
         Min             =   32767
         TabIndex        =   32
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   5040
         TabIndex        =   31
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtTotalPrice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4800
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cboSiding 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtItemCode 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCategoryCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboCategory 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   2415
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   120
         X2              =   6000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label22 
         Caption         =   "Notes"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3600
         TabIndex        =   30
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Total Price"
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Price"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Category Code"
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Siding"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Category "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
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
      Height          =   975
      Left            =   5760
      TabIndex        =   5
      Top             =   1080
      Width           =   6135
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   65
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtContactPerson 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "Cont.Person"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   " Name"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Of Advertisements "
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
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   5535
      Begin MSComctlLib.ListView ListView2 
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
            Picture         =   "frmCreateJobBreif.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateJobBreif.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateJobBreif.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateJobBreif.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateJobBreif.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateJobBreif.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateJobBreif.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
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
   Begin VB.Frame Frame3 
      Caption         =   "List Of Clients"
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
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2778
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
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5760
      X2              =   11760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Brief Date"
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
      Left            =   9000
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Job Brief No"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   720
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
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHow 
         Caption         =   "How to use this System"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmCreateJobBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim billbno, CustomerCode, DateOfQuotation, SelectedQuotationNo, ClientPhysicalAddress, Email

Private Sub cboCouncilPeriod_Click()
        frmCreateJobBrief.ListView1.SetFocus
End Sub

Private Sub cboCountry_Click()
    rmCreateJobBrief.ListView1.SetFocus
End Sub

Private Sub cboElecPeriod_Click()
    frmCreateJobBrief.ListView1.SetFocus
End Sub


Private Sub cboLandLord_Click()
frmCreateJobBrief.ListView1.SetFocus
End Sub

Private Sub cboRentPeriod_Click()
frmCreateJobBrief.ListView1.SetFocus
End Sub



Private Sub cboTown_Click()
frmCreateJobBrief.ListView1.SetFocus
End Sub



Private Sub Combo1_Click()
frmCreateJobBrief.ListView1.SetFocus
End Sub

Private Sub cboCategory_GotFocus()
If Not NewRecord Then Exit Sub
With frmCreateJobBrief
    AttachSQL = "SELECT CategoryName AS SelectField FROM AdvertCategories ORDER BY CategoryName;"
    .cboCategory.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboCity_GotFocus()
If Not NewRecord Then Exit Sub
With frmCreateJobBrief
    AttachSQL = "SELECT Town AS SelectField FROM ODASPTown ORDER BY Town;"
    .cboCity.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboInfoAttached_Click()
frmCreateJobBrief.ListView3.SetFocus
End Sub

Private Sub cboInfoAttached_GotFocus()
If Not NewRecord Then Exit Sub
With frmCreateJobBrief
    AttachSQL = "SELECT InfoName AS SelectField FROM AdvertinfoAttached ORDER BY InfoName;"
    .cboInfoAttached.Clear
    MyCommonData.AttachDropDown
End With
End Sub

Private Sub cboInfoAttached_LostFocus()
''On Error GoTo Err
With frmCreateJobBrief

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertInfoAttached WHERE infoname ='" & Trim(.cboInfoAttached.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboInfoAttached.Text = rsFindRecord!InfoCode & ""
        .dtpDeadlineDate.SetFocus
    End If
    End With
  Exit Sub
err:
     ErrorMessage
End Sub

Private Sub cboSiding_GotFocus()
If Not NewRecord Then Exit Sub
With frmCreateJobBrief
    AttachSQL = "SELECT SidingDescription AS SelectField FROM Advertsiding ORDER BY sidingdescription;"
    .cboSiding.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboCategory_Click()
        frmCreateJobBrief.ListView2.SetFocus
End Sub
Private Sub cboCategory_LostFocus()
'''On Error GoTo Err
With frmCreateJobBrief

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertCategories WHERE CategoryName='" & Trim(.cboCategory.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtCategoryCode.Text = rsFindRecord!categorycode & ""
        .cboCategory.Text = rsFindRecord!categoryname & ""
        
        .ListView2.SetFocus
        
        If .txtCategoryCode.Text = "AAA" Then
            Call ShowAllBillBoardCategories
        Else
            Call ShowBillBoardsPerCategory
        End If
        
    End If
    
End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Function GetAdvertPrice()
'''On Error GoTo Err
With frmCreateJobBrief
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertPricing WHERE BBNo='" & Trim(.txtItemCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        GetAdvertPrice = rsFindRecord!BBCharges & ""
     
        
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Private Sub ShowBillBoardsPerCategory()
'''On Error GoTo Err
With frmCreateJobBrief
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails A ,AdvertCategories B WHERE A.CategoryCode = B.CategoryCode AND A.CategoryCode = '" & Trim(.txtCategoryCode.Text) & "' ORDER BY A.Name;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing:  Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!categorycode))


    If Not IsNull(rsLIST!categoryname) Then
        MyList.SubItems(1) = CStr(rsLIST!categoryname)
    End If
     
    If Not IsNull(rsLIST!BillBoardNo) Then
        MyList.SubItems(2) = CStr(rsLIST!BillBoardNo)
    End If
    
    If Not IsNull(rsLIST!Name) Then
        MyList.SubItems(3) = CStr(rsLIST!Name)
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(4) = CStr(rsLIST!Length)
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(5) = CStr(rsLIST!Width)
    End If
    
    rsLIST.MoveNext
    
Wend


Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Private Sub ShowAllClientsStructure()
'''On Error GoTo Err
With frmCreateJobBrief
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Mobile Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Physical Adress", .ListView1.Width / 5.5

.ListView1.View = lvwReport
End With
Exit Sub
err:
    ErrorMessage
 End Sub
Private Sub ShowAllClients()
'''On Error GoTo Err
With frmCreateJobBrief
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Address", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Mobile Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Physical Adress", .ListView1.Width / 1.5
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients ORDER BY CompanyName;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!customerid))


    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
     
    If Not IsNull(rsLIST!Address) Then
        MyList.SubItems(2) = CStr(rsLIST!Address)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!phone) Then
        MyList.SubItems(4) = CStr(rsLIST!phone)
    End If
    
    If Not IsNull(rsLIST!Mobilephone) Then
        MyList.SubItems(5) = CStr(rsLIST!Mobilephone)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(6) = CStr(rsLIST!Fax)
    End If
    
    If Not IsNull(rsLIST!ContactTitle) Then
        MyList.SubItems(7) = CStr(rsLIST!ContactTitle)
    End If
    
    If Not IsNull(rsLIST!Contactname) Then
        MyList.SubItems(8) = CStr(rsLIST!Contactname)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(9) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!customerid) Then
        MyList.SubItems(10) = CStr(rsLIST!customerid)
    End If
    
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Private Sub ShowAllBillBoardCategoriesStructure()
'''On Error GoTo Err
With frmCreateJobBrief
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 5
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport
End With
Exit Sub
err:
    ErrorMessage
End Sub


Private Sub ShowAllBillBoardCategories()
''On Error GoTo Err
With frmCreateJobBrief
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails A ,AdvertCategories B WHERE A.CategoryCode = B.CategoryCode ORDER BY A.Name;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!categorycode))


    If Not IsNull(rsLIST!categoryname) Then
        MyList.SubItems(1) = CStr(rsLIST!categoryname)
    End If
     
    If Not IsNull(rsLIST!BillBoardNo) Then
        MyList.SubItems(2) = CStr(rsLIST!BillBoardNo)
    End If
    
    If Not IsNull(rsLIST!Name) Then
        MyList.SubItems(3) = CStr(rsLIST!Name)
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(4) = CStr(rsLIST!Length)
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(5) = CStr(rsLIST!Width)
    End If
    
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub cboCity_Click()
frmCreateJobBrief.ListView1.SetFocus
End Sub

Private Sub cboSiding_Click()
frmCreateJobBrief.ListView2.SetFocus
End Sub

Private Sub cboSiding_LostFocus()
''On Error GoTo Err
With frmCreateJobBrief

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertSiding WHERE SidingDescription ='" & Trim(.cboSiding.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboSiding.Text = rsFindRecord!SidingType & ""
        .txtQuantity.SetFocus
    End If
    End With
  Exit Sub
err:
     ErrorMessage
End Sub

Private Sub Check1_Click()

End Sub

Private Sub cboSiteName_Click()
frmCreateJobBrief.ListView3.SetFocus
End Sub

Private Sub cboSiteName_DropDown()
If Not NewRecord Then Exit Sub
With frmCreateJobBrief
    AttachSQL = "SELECT SiteName AS SelectField FROM AdvertSites ORDER BY SiteName;"
    .cboSiteName.Clear
    MyCommonData.AttachDropDown
End With
End Sub

Private Sub cboSiteName_LostFocus()
''On Error GoTo Err
With frmCreateJobBrief

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertSites WHERE Sitename ='" & Trim(.cboSiteName.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboSiteName.Text = rsFindRecord!SiteName & ""
        .txtSideCode.Text = rsFindRecord!SiteNo & ""
        .txtSitePhysicalAddress.Text = rsFindRecord!SitePhysicalAddress & ""
        .txtOtherSite.SetFocus
    End If
    End With
  Exit Sub
err:
     ErrorMessage
End Sub

Private Sub dtpCommenceDate_CloseUp()
''On Error GoTo Err
'With frmCreateJobBrief
'If .dtpCommenceDate.Value < Date Then
'  MsgBox "Commence date can't be in the past", vbExclamation, "Job Brief Commence Date"
'   Exit Sub
'  ElseIf .dtpCommenceDate.Value > .dtpExpectedDOC.Value Then
'  MsgBox "Commence date can not be later than expected date of completion", vbExclamation, "Job Brief Commence Date"
'   Exit Sub
'  ElseIf .dtpCommenceDate.Value > .dtpDeadlineDate.Value Then
'  MsgBox "Commence date can not be later than the deadline date", vbExclamation, "Job Brief Commence Date"
'   Exit Sub
'End If
'End With
'Exit Sub
'Err:
'    ErrorMessage
End Sub

Private Sub dtpDeadlineDate_CloseUp()
''On Error GoTo Err
'With frmCreateJobBrief
'If .dtpDeadlineDate.Value < .dtpExpectedDOC.Value Then
'  MsgBox "Deadlinedate can not be earlier than expected date of completion", vbExclamation, "Job Brief Deadline Date"
'  .dtpDeadlineDate.SetFocus: Exit Sub
'  ElseIf .dtpDeadlineDate.Value < .dtpCommenceDate.Value Then
'  MsgBox "Deadline date can not be earlier than date of commencement", vbExclamation, "Job Brief Deadlinedate"
'  .dtpDeadlineDate.SetFocus: Exit Sub
'  End If
'End With
'Exit Sub
'Err:
'    ErrorMessage

End Sub

Private Sub dtpExpectedDOC_CloseUp()
''On Error GoTo Err
'With frmCreateJobBrief
'If .dtpExpectedDOC.Value > .dtpDeadlineDate Then
'  MsgBox "Expected date of completion can not be later than the deadline date", vbExclamation, "Job Brief Expected D.O.C"
'  .dtpExpectedDOC.SetFocus: Exit Sub
'  ElseIf .dtpExpectedDOC.Value < .dtpCommenceDate.Value Then
'  MsgBox "Expected date of completion can not be earlier than date of commencement", vbExclamation, "Job Brief Expected D.O.C"
'  .dtpExpectedDOC.SetFocus: Exit Sub
'  End If
'End With
'Exit Sub
'Err:
'    ErrorMessage
End Sub

Private Sub Form_Initialize()
Set MyCommonData = New clsCommonData
'Set mycommondata New clsCommonData

End Sub

Private Sub Form_Load()
ShowAllClientsStructure
ShowAllBillBoardCategoriesStructure
End Sub
Private Function AutoPurchaseOrderNo() As String
''On Error GoTo Err
With frmCreateJobBrief

Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(cnCliniclDetails.cnCliniclID) as lastid from cnCliniclDetails"
strLastID = "SELECT MAX(JobBriefNo) AS LastID FROM AdvertJobBrief;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnCOMMON, adOpenKeyset, adLockOptimistic
    If .EOF And .BOF Then 'shows empty recordset
        AutoPurchaseOrderNo = "JB00001" 'format of desired format of the string value
    ElseIf IsNull(!lastid) = True Or !lastid = "" Then
        AutoPurchaseOrderNo = "JB00001"
    Else
       ' If .EOF And .BOF Then .MoveFirst
        '.MoveLast
        strTemp = !lastid
        iNumPos = 1
        Dim sChar As String
        Dim iIDLen As Integer
        iIDLen = Len(strTemp)
        sChar = Mid(strTemp, iNumPos, 1)
        While InStr("1234567890", sChar) = 0
            iNumPos = iNumPos + 1
            sChar = Mid(strTemp, iNumPos, 1)
        Wend
        'store the ID prefix eg AP
        strPrefix = Left(strTemp, iNumPos - 1)
        'store the number portion eg and the length with leading Zeros
        strTemp = Right(strTemp, Len(strTemp) + 1 - iNumPos)
        strTemp = Format(Int(strTemp) + 1, String(iIDLen + 1 - iNumPos, "0"))
        AutoPurchaseOrderNo = strPrefix & strTemp
    End If
End With
End With
    Exit Function
err:
    ErrorMessage
End Function

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''On Error GoTo Err
    If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
    
    If frmCreateJobBrief.ListView1.ListItems.Count = 0 Or frmCreateJobBrief.ListView1.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = frmCreateJobBrief.ListView1.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If frmCreateJobBrief.ListView1.ListItems(i).Text <> Item Then
            frmCreateJobBrief.ListView1.ListItems(i).Checked = False
        End If
    Next i
    
    If frmCreateJobBrief.ListView1.ColumnHeaders(1).Text = "Client Code" Then
      If Item.Checked = True Then
      
        frmCreateJobBrief.txtName.Text = Item.SubItems(1)
        frmCreateJobBrief.txtAddress.Text = Item.SubItems(2)
        frmCreateJobBrief.cboCity.Text = Item.SubItems(3)
        frmCreateJobBrief.txtTelephone.Text = Item.SubItems(4)
        frmCreateJobBrief.txtMobile.Text = Item.SubItems(5)
        frmCreateJobBrief.txtFax.Text = Item.SubItems(6)
        frmCreateJobBrief.txtContactPerson.Text = Item.SubItems(8)
        frmCreateJobBrief.txtContactTitle.Text = Item.SubItems(7)
        CustomerCode = Item.SubItems(10)
             
    ElseIf Item.Checked = False Then
    End If
   ElseIf frmCreateJobBrief.ListView1.ColumnHeaders(1).Text = "Quotation Number" Then
        If Item.Checked = True Then
        
        frmCreateJobBrief.txtQuotationNo.Text = Item
        SelectedQuotationNo = Item
        DateOfQuotation = Item.SubItems(1)
        frmCreateJobBrief.txtName = Item.SubItems(2)
        frmCreateJobBrief.txtContactPerson.Text = Item.SubItems(3)
        frmCreateJobBrief.txtContactTitle.Text = Item.SubItems(4)
        ClientPhysicalAddress = Item.SubItems(5)
        frmCreateJobBrief.txtMobile = Item.SubItems(6)
        frmCreateJobBrief.txtTelephone.Text = Item.SubItems(7)
        Email = Item.SubItems(8)
        frmCreateJobBrief.txtFax = Item.SubItems(9)
        frmCreateJobBrief.txtAddress.Text = Item.SubItems(10)
        Call ShowAdverisementsUnderSelectedQuotation
      
     ElseIf Item.Checked = False Then
     End If
 End If
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ShowAdverisementsUnderSelectedQuotation()
''On Error GoTo Err
With frmCreateJobBrief
.ListView3.ListItems.Clear
.ListView3.ColumnHeaders.Clear

.ListView3.ColumnHeaders.Add , , "AdvertCode", .ListView3.Width / 6#
.ListView3.ColumnHeaders.Add , , "Quotatin No", .ListView3.Width / 4
.ListView3.ColumnHeaders.Add , , "Advert Name", .ListView3.Width / 5.5
.ListView3.ColumnHeaders.Add , , "Media Name", .ListView3.Width / 4
.ListView3.ColumnHeaders.Add , , "Media Code", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Client", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Illuminated", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Siding", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Bordered", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Quantity", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Price", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Total Price", .ListView3.Width / 6.5

.ListView3.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertQuotation A ,AdvertQuotationItems B,AdvertSiding C WHERE A.QuotationNo = B.QuotationNo AND B.SidingType = C.SidingType AND B.QuotationNo = '" & SelectedQuotationNo & "' ORDER BY B.ItemCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView3.View = lvwList
    Set MyList = .ListView3.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!ItemCode))


    If Not IsNull(rsLIST!QuotationNo) Then
        MyList.SubItems(1) = CStr(rsLIST!QuotationNo)
    End If
     
    If Not IsNull(rsLIST!ItemName) Then
        MyList.SubItems(2) = CStr(rsLIST!ItemName)
    End If
    
    If Not IsNull(rsLIST!categoryname) Then
        MyList.SubItems(3) = CStr(rsLIST!categoryname)
    End If
    
    If Not IsNull(rsLIST!categorycode) Then
        MyList.SubItems(4) = CStr(rsLIST!categorycode)
    End If
    
    If Not IsNull(rsLIST!QuotationDesc) Then
        MyList.SubItems(5) = CStr(rsLIST!QuotationDesc)
    End If
    
    If Not IsNull(rsLIST!illuminated) And (rsLIST!illuminated) = 0 Then
        MyList.SubItems(6) = CStr("NO")
      ElseIf Not IsNull(rsLIST!illuminated) And (rsLIST!illuminated) = 1 Then
        MyList.SubItems(6) = CStr("YES")
    End If
    
    If Not IsNull(rsLIST!SidingDescription) Then
        MyList.SubItems(7) = CStr(rsLIST!SidingDescription)
    End If
    
    
    If Not IsNull(rsLIST!BorderType) And (rsLIST!BorderType) = 0 Then
        MyList.SubItems(8) = CStr("NO")
      ElseIf Not IsNull(rsLIST!BorderType) And (rsLIST!BorderType) = 1 Then
        MyList.SubItems(8) = CStr("YES")
    End If
    
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(9) = CStr(rsLIST!quantity)
    End If
    
    If Not IsNull(rsLIST!Price) Then
        MyList.SubItems(10) = FormatNumber((rsLIST!Price), 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalPrice) Then
        MyList.SubItems(11) = FormatNumber((rsLIST!TotalPrice), 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If frmCreateJobBrief.ListView2.ListItems.Count = 0 Or frmCreateJobBrief.ListView2.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = frmCreateJobBrief.ListView2.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If frmCreateJobBrief.ListView2.ListItems(i).Text <> Item Then
            frmCreateJobBrief.ListView2.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then
    
        
        frmCreateJobBrief.txtCategoryCode.Text = Item
        frmCreateJobBrief.cboCategory.Text = Item.SubItems(1)
        frmCreateJobBrief.txtItemCode.Text = Item.SubItems(2)
        frmCreateJobBrief.txtItemName.Text = Item.SubItems(3)
        frmCreateJobBrief.txtPrice.Text = GetAdvertPrice
        
        
        
    ElseIf Item.Checked = False Then
    
        
    End If
    
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub ClearTextFields()
For Each i In Screen.ActiveForm
    If TypeOf i Is TextBox And i.Name <> "txtTitle" Then
        i.Text = Empty
    End If
    If TypeOf i Is ComboBox Then
        i.Clear
    End If
    If TypeOf i Is Image Then
        i.Picture = LoadPicture("")
    End If
Next i
End Sub

Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If frmCreateJobBrief.ListView3.ListItems.Count = 0 Or frmCreateJobBrief.ListView3.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = frmCreateJobBrief.ListView3.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If frmCreateJobBrief.ListView3.ListItems(i).Text <> Item Then
            frmCreateJobBrief.ListView3.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then
    
        
        frmCreateJobBrief.txtItemCode.Text = Item
        frmCreateJobBrief.txtItemName.Text = Item.SubItems(2)
        frmCreateJobBrief.cboCategory.Text = Item.SubItems(3)
        frmCreateJobBrief.txtCategoryCode.Text = Item.SubItems(4)
        frmCreateJobBrief.cboSiding.Text = Item.SubItems(7)
        frmCreateJobBrief.txtQuantity.Text = Item.SubItems(9)
        frmCreateJobBrief.txtPrice.Text = Item.SubItems(10)
        frmCreateJobBrief.txtTotalPrice.Text = Item.SubItems(11)
        
        If Item.SubItems(6) = "YES" Then
        frmCreateJobBrief.chkIlluminate.Value = 1
        Else
        frmCreateJobBrief.chkIlluminate.Value = 0
        End If
        
        If Item.SubItems(8) = "YES" Then
        frmCreateJobBrief.chkBorder.Value = 1
        Else
        frmCreateJobBrief.chkBorder.Value = 0
        End If
        
        
    ElseIf Item.Checked = False Then
    
        
    End If
    
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuClear_Click()
    MyCommonData.ClearTextFields
End Sub

Private Sub mnuClose_Click()
    Unload frmCreateJobBrief
End Sub

Private Sub mnuCurrent_Click()
    Call ShowCurrentSettings

End Sub
Private Sub ShowCurrentSettings()
'''On Error GoTo Err
With frmCreateJobBrief
Screen.MousePointer = vbHourglass

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Brokers Code", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Brokers Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Branch", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Town/City", .ListView1.Width / 10
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Phone Number", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Email", .ListView1.Width / 5

.ListView1.View = lvwReport: .ListView1.Visible = True

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ParamInsuranceBrokers.* FROM ParamInsuranceBrokers ORDER BY ParamInsuranceBrokers.BrokersCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem, NCount As Double

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Screen.MousePointer = vbDefault: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BrokersCode))
    
    If Not IsNull(rsLIST!BrokersCode) Then
        MyList.SubItems(1) = CStr(rsLIST!BrokersName)
        MyList.SubItems(2) = CStr(rsLIST!PhysicalAddress)
        MyList.SubItems(3) = CStr(rsLIST!postaladdress)
        MyList.SubItems(4) = CStr(rsLIST!Branch)
        MyList.SubItems(5) = CStr(rsLIST!towncity)
        MyList.SubItems(6) = CStr(rsLIST!country)
        MyList.SubItems(7) = CStr(rsLIST!ContactPerson)
        MyList.SubItems(8) = CStr(rsLIST!ContactTitle)
        MyList.SubItems(9) = CStr(rsLIST!TelephoneNo)
        MyList.SubItems(10) = CStr(rsLIST!Email)

    End If

    rsLIST.MoveNext
    
Wend

    Set MyList = Nothing: Set rsLIST = Nothing: Screen.MousePointer = vbDefault
        
End With
Exit Sub
err:
If err.Number = 3265 Then
    Resume Next
Else
    Screen.MousePointer = vbDefault
    ErrorMessage
End If
End Sub
Private Function ValidRecord() As Boolean
With frmCreateJobBrief
    If .txtContactPerson.Text = Empty Then
        strMessage = "Name Of Contact Person Required...!"
        .txtContactPerson.SetFocus
    ElseIf .txtName.Text = Empty Then
        strMessage = "Clent name required...!"
        .txtName.SetFocus
    ElseIf .txtItemCode.Text = Empty Then
        strMessage = "BillBoard number required...!"
        .txtItemCode.SetFocus
    ElseIf .cboSiding.Text = Empty Then
        strMessage = "Advertisement siding required...!"
        .cboSiding.SetFocus
    ElseIf .txtQuantity.Text = Empty Then
        strMessage = "Quantity Required...!"
        .txtQuantity.SetFocus
    ElseIf .txtPrice.Text = Empty Then
        strMessage = "Advert Price Required...!"
        .txtPrice.SetFocus
    ElseIf .txtTotalPrice.Text = Empty Then
        strMessage = "Total Price Required...!"
        .txtTotalPrice.SetFocus
     ElseIf .txtJobBriefNo.Text = Empty Then
        strMessage = "JobBrief Number required...!"
        .txtJobBriefNo.SetFocus
    
    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function
Private Function ValidMainRecord() As Boolean
With frmCreateJobBrief
    If .txtJobBriefNo.Text = Empty Then
        strMessage = "Job Brief Number required...!"
        .txtJobBriefNo.SetFocus
    ElseIf .txtName.Text = Empty Then
        strMessage = "Clent name required...!"
        .txtName.SetFocus
    ElseIf .dtpBriefDate.Value = Empty Then
        strMessage = "Quotation date required...!"
        .dtpBriefDate.SetFocus
     
    Else
        ValidMainRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function
Public Sub RemoveCurrentList2Item()
''On Error GoTo Err
With frmCreateJobBrief
Dim i, j, k
   j = .ListView2.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView2.ListItems(i).Checked = True Then
         .ListView2.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
err:
   ErrorMessage
End Sub
Public Sub RemoveCurrentList3Item()
''On Error GoTo Err
With frmCreateJobBrief
Dim i, j, k
   j = .ListView3.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView3.ListItems(i).Checked = True Then
         .ListView3.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
err:
   ErrorMessage
End Sub
Private Sub mnuRegisteredClients_Click()
If Not NewRecord Then Exit Sub
Call ShowAllClients
End Sub

Private Sub mnuShowQuotations_Click()
        If Not NewRecord Then Exit Sub
        Call ShowAllUnAuthorizedQuotations
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''On Error GoTo Err
    Dim TotalSum, TotalQuantity, BriefDate, CommenceDate, DOCDate, DeadDate As Variant
    With frmCreateJobBrief
    
    BriefDate = Format(.dtpBriefDate.Value, "MMMM dd,yyyy")
    CommenceDate = Format(.dtpCommenceDate.Value, "MMMM dd,yyyy")
    DOCDate = Format(.dtpExpectedDOC.Value, "MMMM dd,yyyy")
    DeadDate = Format(.dtpDeadlineDate.Value, "MMMM dd,yyyy")
    
    Select Case Button.Key
    Case "N"
    Select Case Button.Caption
    Case "New &Record "
        If EditRecord Then Exit Sub
        MyCommonData.ClearTextFields: .ListView3.ListItems.Clear: .ListView1.ListItems.Clear: .ListView2.ListItems.Clear: .txtName.SetFocus
        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
        .txtName.SetFocus
        .txtJobBriefNo.Text = AutoPurchaseOrderNo
        .dtpBriefDate.Value = MyCurrentDate
        .dtpCommenceDate.Value = MyCurrentDate
        .dtpExpectedDOC.Value = DateAdd("d", CDbl(3), MyCurrentDate)
        .dtpDeadlineDate.Value = DateAdd("d", CDbl(3), .dtpExpectedDOC.Value)
        
    Case "&Save Record "
        If NewRecord Then
        If ValidRecord Then
            NewSQL = "INSERT INTO AdvertJobBriefItems(JobBriefNo,QuotationNo,ItemCode,ItemName,CategoryCode,CategoryName,ClientName,Illuminate,Siding,Border,Quantity,Price,TotalPrice,LPONo,InfoAttached,Comments,SiteCode,OtherSite,CreatedBy,DateCreated,AccPeriod)VALUES('" & Trim(.txtJobBriefNo.Text) & "','" & Trim(.txtQuotationNo.Text) & "','" & Trim(.txtItemCode.Text) & "','" & Trim(.txtItemName.Text) & "','" & Trim(.txtCategoryCode.Text) & "','" & Trim(.cboCategory.Text) & "','" & Trim(.txtName.Text) & "','" & .chkIlluminate.Value & _
            "','" & Trim(.cboSiding.Text) & "','" & .chkBorder.Value & "','" & Trim(.txtQuantity.Text) & "'," & CCur(.txtPrice.Text) & "," & CCur(.txtTotalPrice.Text) & ",'" & Trim(.txtLpono.Text) & "','" & Trim(.cboInfoAttached.Text) & "','" & Trim(.txtJobBriefComments.Text) & "','" & Trim(.txtSideCode.Text) & "','" & Trim(.txtOtherSite.Text) & "','" & Trim(CurrentUserName) & "','" & Trim(MyCurrentDate) & "','" & Trim(MyCurrentPeriod) & "');"
            Set rsNewRecord = New ADODB.Recordset
            rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsNewRecord = Nothing
            RemoveCurrentList3Item
            RemoveCurrentList2Item
             Button.Caption = "NE&XT ITEM"
            .Toolbar1.Buttons(3).Caption = "FINISH"
            
        End If
        End If
        
     Case "NE&XT ITEM"
          .txtItemCode.Text = ""
          .txtItemName.Text = ""
          .cboCategory.Text = ""
          .txtCategoryCode.Text = ""
          .chkIlluminate.Value = 0
          .cboSiding.Text = ""
          .chkBorder.Value = 0
          .txtQuantity.Text = ""
          .txtPrice.Text = ""
          .txtTotalPrice = ""
          .cboSiteName.Text = ""
          .txtSitePhysicalAddress.Text = ""
          .txtOtherSite.Text = ""
          .txtSideCode.Text = ""
          .chkSiteAvailable(0).Value = 0
          .chkSiteAvailable(1).Value = 0
          .chkReserve.Value = 0
          .cboInfoAttached.Text = ""
          RemoveCurrentList3Item
          RemoveCurrentList2Item
          Button.Caption = "&Save Record ": Button.Image = 4
    Case Else
        Exit Sub
    End Select
    
Case "E"
    Select Case Button.Caption
'    Case "&Edit/Change "
'    If NewRecord Then Exit Sub
'        If .txtCode.Text = Empty Then
'            MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
'           .txtCode.SetFocus
'        Else
'           .txtCode.Locked = True
'            Button.Caption = "Save &Changes ": Button.Image = 4
'            EditRecord = True
'        End If
'    Case "Save &Changes "
'        If EditRecord Then
'        If ValidRecord Then
'        EditSQL = "Update ParamInsuranceBrokers SET BrokersName = '" & Trim(txtName.Text) & "'" & _
'        " ,Branch = '" & Trim(txtBranch.Text) & "'" & _
'        " ,PhysicalAddress = '" & Trim(txtAddress1.Text) & "'" & _
'        " ,PostalAddress = '" & Trim(txtAddress2.Text) & "'" & _
'        " ,TownCity = '" & Trim(cboTown.Text) & "'" & _
'        " ,Country = '" & Trim(cboCountry.Text) & "'" & _
'        " ,ContactPerson = '" & Trim(txtPerson.Text) & "'" & _
'        " ,ContactTitle = '" & Trim(cboTitle.Text) & "'" & _
'        " ,TelephoneNo = '" & Trim(txtPhone.Text) & "'" & _
'        " ,Email = '" & Trim(txtEmail.Text) & "' WHERE BrokersCode='" & Trim(txtCode.Text) & "';"
'
'            Set rsEditRecord = New ADODB.Recordset
'            rsEditRecord.Open EditSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'            Set rsEditRecord = Nothing
'            .txtCode.Locked = False: EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 5
'        End If
'        End If
     Case "FINISH"
         If ValidMainRecord Then
           
            Set rsNewRecord = New ADODB.Recordset
            rsNewRecord.Open "INSERT INTO AdvertJobBrief(LPONo,JobBriefNo,ExpectedDOC,QuotationNumber,QuotationDate,BriefDate,CustomerNumber,CustomerName,ContactTitle,ContactName,PhysicalAddress,MobilePhone,Phone,Email,Fax,CommenceDate,DeadlineDate,InfoAttached,DateCreated,CreatedBy,AccPeriod)VALUES('" & Trim(.txtLpono.Text) & "','" & Trim(.txtJobBriefNo.Text) & "','" & DOCDate & "','" & SelectedQuotationNo & "','" & DateOfQuotation & "','" & BriefDate & "','" & CustomerCode & "','" & Trim(.txtName.Text) & "','" & Trim(.txtContactTitle.Text) & "','" & Trim(.txtContactPerson.Text) & "','" & ClientPhysicalAddress & "','" & Trim(.txtMobile.Text) & "','" & Trim(.txtTelephone.Text) & "','" & Email & "','" & Trim(.txtFax.Text) & "','" & CommenceDate & "','" & DeadDate & "','" & Trim(.cboInfoAttached.Text) & "','" & MyCurrentDate & "','" & CurrentUserName & "','" & MyCurrentPeriod & "')", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsNewRecord = Nothing
         
            Set rsFindRecord = New ADODB.Recordset
               rsFindRecord.Open "SELECT SUM(TotalPrice)as Total FROM AdvertJobBriefItems WHERE JobBriefNo = '" & Trim(.txtJobBriefNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
               If rsFindRecord.EOF And rsFindRecord.BOF Then Set rsFindRecord = Nothing
               If Not IsNull(rsFindRecord!Total) Then
               TotalSum = CCur(rsFindRecord!Total)
               Else
                TotalSum = 0
               End If
               End If
             Set rsFindRecord = Nothing
               
            Set rsLineUpdate = New ADODB.Recordset
               rsLineUpdate.Open "UPDATE AdvertJobBrief SET TotalPrice = " & CCur(TotalSum) & " WHERE JobBriefNo = '" & Trim(.txtJobBriefNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsLineUpdate = Nothing
            
            Set rsLineUpdate = New ADODB.Recordset
               rsLineUpdate.Open "UPDATE AdvertQuotation SET JobBriefStatus = '" & "Y" & "' WHERE QuotationNo = '" & Trim(.txtQuotationNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsLineUpdate = Nothing
            
       .Toolbar1.Buttons(2).Caption = "New &Record "
       .Toolbar1.Buttons(2).Image = 2
       .Toolbar1.Buttons(3).Image = 5
       .Toolbar1.Buttons(3).Caption = "&Edit/Change "
             
'     End If
    Case Else
       
        Exit Sub
    End Select
Case "S"
'    If NewRecord Or EditRecord Then Exit Sub
'    INPQRY = InputBox("Please Enter the Brokers Code for the Record to Search and Display...!!!", "Enter Brokers Code...")
'    If Len(INPQRY) = 0 Then
'        MsgBox "Required Search Parameter Missing or the Operation Was Cancelled...! No Work was Done!!!", vbCritical + vbOKOnly, "Missing Parameter"
'        Exit Sub
'    Else
'        Set rsFindRecord = cnCOMMON.Execute("SELECT ParamInsuranceBrokers.* FROM ParamInsuranceBrokers WHERE ParamInsuranceBrokers.BrokersCode='" & Trim(INPQRY) & "';")
'        If rsFindRecord.EOF And rsFindRecord.BOF Then
'            MsgBox "Requested Record Missing or Has Been Deleted. Check your Entries to Ensure they are Accurately Spelt...!", vbOKOnly + vbExclamation, "Record NOT Found...!"
'            Set rsFindRecord = Nothing: Exit Sub
'        Else
'            .txtCode.Text = Trim(rsFindRecord!BrokersCode & "")
'            .txtName.Text = Trim(rsFindRecord!BrokersName & "")
'            .txtBranch.Text = Trim(rsFindRecord!Branch & "")
'            .txtAddress1.Text = Trim(rsFindRecord!PhysicalAddress & "")
'            .txtAddress2.Text = Trim(rsFindRecord!PostalAddress & "")
'            .cboTown.Text = Trim(rsFindRecord!TownCity & "")
'            .cboCountry.Text = Trim(rsFindRecord!Country & "")
'            .txtPerson.Text = Trim(rsFindRecord!ContactPerson & "")
'            .cboTitle.Text = Trim(rsFindRecord!ContactTitle & "")
'            .txtPhone.Text = Trim(rsFindRecord!TelephoneNo & "")
'            .txtEmail.Text = Trim(rsFindRecord!Email & "")
'
'        End If
'        Set rsFindRecord = Nothing
'    End If
Case "R"
    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
        .Toolbar1.Buttons(2).Caption = "New &Record "
        .Toolbar1.Buttons(2).Image = 2
        .Toolbar1.Buttons(3).Caption = "&Edit/Change "
        .Toolbar1.Buttons(3).Image = 5
        NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
Case "P"
'    Load frmRptAdvertPrintOut
'    frmRptAdvertPrintOut.Show 1, frmCreateJobBrief
Case "F"
     
     
Case Else
    Exit Sub
End Select
End With
Exit Sub
err:
    ErrorMessage

End Sub




Private Sub txtSidingCost_Change()

End Sub


Private Sub txtQuantity_Change()
'On Error GoTo err
With frmCreateJobBrief
.txtTotalPrice.Text = Val(.txtQuantity) * Val(.txtPrice.Text)
End With
Exit Sub
err:
   ErrorMessage
End Sub

Private Sub VScroll1_Change()
With frmCreateJobBrief
        .txtQuantity.Text = .VScroll1.Value
End With
End Sub

Private Sub VScroll1_GotFocus()
With frmCreateJobBrief
.VScroll1.Value = 1
End With
End Sub
