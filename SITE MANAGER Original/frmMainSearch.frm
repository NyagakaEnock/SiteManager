VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRQuotationList 
   Caption         =   "Print Quotation Listing for the Selected Period"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.OptionButton optNotConverted 
         Caption         =   "Not Converted"
         Height          =   375
         Left            =   6000
         TabIndex        =   16
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optDispatched 
         Caption         =   "Dispatched?"
         Height          =   255
         Left            =   9600
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optNotPrinted 
         Caption         =   "Not Printed?"
         Height          =   375
         Left            =   8520
         TabIndex        =   14
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton optPrinted 
         Caption         =   "Printed"
         Height          =   255
         Left            =   7320
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optConverted 
         Caption         =   "Converted"
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate Data"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   5520
         Width           =   3495
      End
      Begin VB.Frame Frame8 
         Height          =   4815
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   10695
         Begin VB.TextBox txtReportTitle 
            Height          =   315
            Left            =   3120
            TabIndex        =   17
            Top             =   240
            Width           =   6735
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4095
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   7223
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label3 
            Caption         =   "Report Title"
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   270
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdPRINT 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   7
         Top             =   5520
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker DTPickerLastDate 
         Height          =   315
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57409537
         CurrentDate     =   39255
      End
      Begin MSComCtl2.DTPicker DTPickerStartDate 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57409537
         CurrentDate     =   39255
      End
      Begin VB.TextBox txtLastDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtStartDate 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   0
         TabIndex        =   11
         Top             =   6120
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Last Date"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRQuotationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckDetails()
'On Error GoTo err
    With frmRQuotationList
            If CDate(.txtStartDate.Text) > CDate(.txtLastDate.Text) Then
                    MsgBox "The Last Date Must be After the Start Date"
                    .txtStartDate.SetFocus
            
            ElseIf .txtLastDate.Text <= "" Then
                    MsgBox "The Last Date Cannot be Blank............."
                    .txtLastDate.SetFocus
            
            ElseIf .txtStartDate.Text <= "" Then
                    MsgBox "The Start Date Cannot be Blank............."
                    .txtStartDate.SetFocus
            
            Else
                    bSaveRECORD = True
                    Me.cmdPrint.Enabled = True
            End If
    
    End With

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdPrint_Click()
        If bSaveRECORD = True Then
                Load frmRQuotationListing
                frmRQuotationListing.Show 1, Me
        End If
End Sub

Private Sub Command1_Click()
        bSaveRECORD = False
        CheckDetails
        If bSaveRECORD = True Then showALLQuotationListings
End Sub

Private Sub DTPickerStartDate_CloseUp()
        Me.txtStartDate.Text = Me.DTPickerStartDate.Value
End Sub
Private Sub DTPickerLastDate_CloseUp()
        Me.txtStartDate.Text = Me.DTPickerLastDate.Value
End Sub

Private Sub Form_Activate()
        DisableRECORD
        loadRECORD
        showALLQuotationListings
        
        If bSaveRECORD = True Then
                Me.cmdPrint.Enabled = True
        Else: Me.cmdPrint.Enabled = False
        End If
        
End Sub
Private Sub loadRECORD()
'On Error GoTo err
    With frmRQuotationList
            .txtLastDate.Text = Date
            .txtStartDate.Text = Date
            .DTPickerLastDate.Value = Date
            .DTPickerStartDate.Value = Date
            .txtReportTitle.Text = "QUOTATION REGISTER"
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub DisableRECORD()
'On Error GoTo err
    With frmRQuotationList
            .txtLastDate.Locked = True
            .txtStartDate.Locked = True
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub optConverted_Click()
        Me.txtReportTitle.Text = "QUOTATION LISTING - DISPATCHED"
End Sub

Private Sub optDispatched_Click()
        Me.txtReportTitle.Text = "QUOTATION LISTING - PRINTED"
End Sub

Private Sub optNotConverted_Click()
        Me.txtReportTitle.Text = "QUOTATION LISTING - NOT CONVERTED"
End Sub

Private Sub optNotPrinted_Click()
        Me.txtReportTitle.Text = "QUOTATION LISTING - NOT PRINTED"
End Sub

Private Sub optPrinted_Click()
        Me.txtReportTitle.Text = "QUOTATION LISTING - PRINTED"
End Sub
