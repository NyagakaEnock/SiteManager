VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASMCouncilRates 
   Caption         =   "COUNCIL RATES"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   1470
   ClientWidth     =   11685
   Icon            =   "frmODASMCouncilRates.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASMCouncilRates.frx":0442
   ScaleHeight     =   7380
   ScaleWidth      =   11685
   Begin VB.Frame Frame6 
      Caption         =   "Council Accounts"
      Height          =   2295
      Left            =   120
      TabIndex        =   46
      Top             =   5040
      Width           =   5295
      Begin MSComctlLib.ListView ListView5 
         Height          =   1935
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
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
   Begin VB.Frame Frame5 
      Caption         =   "Job Briefs"
      Height          =   2175
      Left            =   5520
      TabIndex        =   43
      Top             =   2880
      Width           =   6135
      Begin VB.TextBox txtJBDurationMode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3000
         TabIndex        =   58
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox txtJBExpiryDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   55
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtJBStartDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   54
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtJBDuration 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   53
         Top             =   1680
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   1335
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
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
      Begin VB.Label Label21 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   57
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1710
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sites With Rates"
      Height          =   2175
      Left            =   120
      TabIndex        =   39
      Top             =   2880
      Width           =   5295
      Begin MSComctlLib.ListView ListView3 
         Height          =   1815
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3201
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
      Caption         =   "Rates Details"
      Height          =   2295
      Left            =   5520
      TabIndex        =   23
      Top             =   5040
      Width           =   6135
      Begin VB.TextBox txtAccountNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   51
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtJobBriefItemNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   50
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtReferenceNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4200
         TabIndex        =   41
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtRateStartDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   30
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   29
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox txtRateDueDate 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4200
         TabIndex        =   28
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtCurrentYear 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   27
         Top             =   1410
         Width           =   1335
      End
      Begin VB.TextBox txtRateExpiryDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   26
         Top             =   630
         Width           =   1815
      End
      Begin VB.ComboBox cboPaymentMode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   3120
         TabIndex        =   25
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox txtPaymentDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   24
         Top             =   1410
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPickerRateStartDate 
         Height          =   315
         Left            =   2160
         TabIndex        =   31
         Top             =   630
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   68812801
         CurrentDate     =   38400
      End
      Begin VB.Label Label16 
         Caption         =   "Account No"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Job Brief No"
         Height          =   255
         Left            =   2520
         TabIndex        =   49
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Reference No"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Status"
         Height          =   255
         Left            =   2520
         TabIndex        =   42
         Top             =   1830
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Mode"
         Height          =   255
         Left            =   2520
         TabIndex        =   37
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   " Due Date"
         Height          =   255
         Left            =   2520
         TabIndex        =   35
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Year"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Expiry"
         Height          =   255
         Left            =   3720
         TabIndex        =   33
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Last Payment Date"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sites Without Rates"
      Height          =   2175
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   5295
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3201
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
   Begin VB.Frame Frame12 
      Caption         =   "Site Details"
      Height          =   2175
      Left            =   5520
      TabIndex        =   0
      Top             =   720
      Width           =   6135
      Begin VB.TextBox txtSiteDetails 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   22
         Top             =   1386
         Width           =   3735
      End
      Begin VB.TextBox txtSiteNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   19
         Top             =   1386
         Width           =   1215
      End
      Begin VB.TextBox txtMediaCode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   1004
         Width           =   1215
      End
      Begin VB.TextBox txtMediaSize 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3240
         TabIndex        =   16
         Top             =   1004
         Width           =   2655
      End
      Begin VB.TextBox txtDuration 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   12
         Top             =   1770
         Width           =   375
      End
      Begin VB.TextBox txtTown 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtTownCode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCommencementDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   1770
         Width           =   1215
      End
      Begin VB.TextBox txtExpiryDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   4
         Top             =   1770
         Width           =   1695
      End
      Begin VB.TextBox txtPlotNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   622
         Width           =   1215
      End
      Begin VB.TextBox txtPlotName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   622
         Width           =   3735
      End
      Begin VB.Label Label12 
         Caption         =   "Site No"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1416
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Media Size"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   1034
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Media"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1034
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "M"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Town"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Plot"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   652
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
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
      Begin VB.TextBox txtMast 
         Height          =   285
         Left            =   11040
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10800
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
               Picture         =   "frmODASMCouncilRates.frx":0784
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMCouncilRates.frx":0DFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMCouncilRates.frx":1340
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMCouncilRates.frx":1792
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMCouncilRates.frx":1AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMCouncilRates.frx":2126
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMCouncilRates.frx":27A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMCouncilRates.frx":2BF2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10920
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Menu mnuViewRateSchedule 
      Caption         =   "View Listings"
      Begin VB.Menu mnuViewCouncilRates 
         Caption         =   "View Council Rates"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSchedule 
         Caption         =   "View Schedule"
      End
   End
End
Attribute VB_Name = "frmODASMCouncilRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSEND As clsODASCouncilRates, MyCommonData As clsCommonData

Private Sub cboPaymentMode_GotFocus()
    SelectModeGotFocus
End Sub

Private Sub cboPaymentMode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboPaymentMode_LostFocus()
    selectModeLostFocus
End Sub

Private Sub DTPickerRateStartDate_CloseUp()
    Set rsSEND = New clsODASCouncilRates
    rsSEND.calcDUEDATE
    Set rsSEND = Nothing

End Sub

Private Sub Form_Activate()
        Set MyCommonData = New clsCommonData
        Set rsSEND = New clsODASCouncilRates
        disableALLRECORD
        ShowSITESWITHRATES
        ListALLSITES
        ListALLCOUNCILACCOUNTS
        Me.txtRateStartDate.Text = Date
        frmODASMCouncilRates.Frame2.Enabled = False
End Sub

Private Sub Form_Terminate()
        Set rsSEND = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

        If NewRecord Then
            Cancel = True
            MsgBox "System cannot close this form when the generation of council rates is in Process...", vbCritical + vbOKOnly, "Interruption of Job Brief creation process"
        Else
            Cancel = False
            Set rsSEND = Nothing
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
            
            frmODASMCouncilRates.txtSiteNo.Text = Item.Text
            frmODASMCouncilRates.txtPlotName.Text = Item.SubItems(2)
            frmODASMCouncilRates.txtMast.Text = Item.SubItems(1)
            Set rsSEND = New clsODASCouncilRates
            rsSEND.loadRECORD
            Set rsSEND = Nothing
            
        Else
            Item.Checked = False
        End If
        
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView4.SortKey = ColumnHeader.Index - 1
    ListView4.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView4_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView4.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView4.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView4.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMCouncilRates.txtJobBriefItemNo.Text = Item.Text
            Set rsSEND = New clsODASCouncilRates
            rsSEND.loadJOBBRIEFITEMS
            rsSEND.loadSTARTDATE
            rsSEND.loadCOUNCILRATES
            rsSEND.calcDUEDATE

            Set rsSEND = Nothing

        Else
            Item.Checked = False
        End If
        
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView5.SortKey = ColumnHeader.Index - 1
    ListView5.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView5_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView5.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView5.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView5.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMCouncilRates.txtAccountNo.Text = Item.Text
            
        Else
            Item.Checked = False
        End If
        
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuViewCouncilRates_Click()
On Error GoTo err
        Load frmODASVCouncilRates
        frmODASVCouncilRates.Show 1, Me
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuViewSchedule_Click()
On Error GoTo err
        Load frmODASMRateSchedule
        frmODASMRateSchedule.Show 1, Me
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Set rsSEND = New clsODASCouncilRates

        Select Case Button.Key
            Case "N"
                enableALLRECORD
                NewRecord = True
                frmODASMCouncilRates.Frame2.Enabled = True

                Select Case Button.Caption
                    Case "New &Record "
                            enableALLRECORD
                            Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Save Record"
                            rsSEND.ValidateRECORD
                            If bSaveRECORD = True Then
                                    rsSEND.saveRecord
                                    rsSEND.updateSITE
                                    ShowSITESWITHRATES
                                    ListALLSITES

                                    disableALLRECORD
                                    NewRecord = False: Button.Caption = "New &Record ": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                            End If
                End Select
            Case "E"
                 Select Case Button.Caption
                    Case "Edit &Change "
                         If NewRecord Then Exit Sub
                              
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                EditRecord = True
                    Case "Save &Change "
                            If bSaveRECORD = True Then
                                Button.Caption = "Edit &Change ": Button.Image = 6
                            End If
                End Select
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                    Case "Cancel"
                            clearALLRECORD
                            disableALLRECORD
                End Select
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 3
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 6
                    .Toolbar1.Buttons(4).Caption = "&Search/Find "
                    .Toolbar1.Buttons(4).Image = 7
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen

            Case "H"
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 12
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            Case "P"
                CouncilForm = True
                    INPQRY2 = .txtCurrentYear.Text
                    If (bBillBoard = True Or bStreetSign = True) Then
                        CurrentRecord = .txtMast.Text
                    Else
                        CurrentRecord = .txtSiteNo.Text
                    End If
                    Load frmODASRRatesSchedule
                    frmODASRRatesSchedule.Show vbModal
        End Select
                    
        Set rsSEND = Nothing

    End With
Exit Sub
err:
ErrorMessage
End Sub

