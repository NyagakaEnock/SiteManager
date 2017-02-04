VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmODASMLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Brief Ledger"
   ClientHeight    =   5715
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11445
   Icon            =   "frmODASMLedger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Invoices"
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
      Height          =   2175
      Left            =   5400
      TabIndex        =   17
      Top             =   3480
      Width           =   6015
      Begin VB.TextBox txtTotalInvoices 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   2295
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label6 
         Caption         =   "Total Invoices"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "All Receipts Received"
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
      Height          =   2775
      Left            =   5400
      TabIndex        =   15
      Top             =   720
      Width           =   6015
      Begin VB.TextBox txtTotalReceipts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         MaxLength       =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   2415
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3625
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
      Begin VB.Label Label5 
         Caption         =   "Total"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.TextBox txtJobBriefDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame FrameSites 
      Caption         =   " Job Brief items"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   5175
      Begin MSComctlLib.ListView ListView5 
         Height          =   1815
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
   Begin VB.TextBox txtJobBriefNo 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1335
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
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtTotalCostInclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3600
         TabIndex        =   24
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   2400
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtTotalCostExclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   22
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   13
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         MaxLength       =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "TOTAL         Cost Exclusive           VAT              Cost Incl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label9 
         Caption         =   "Company"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Product"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Description"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   " Name"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
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
            Picture         =   "frmODASMLedger.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLedger.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLedger.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLedger.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLedger.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLedger.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLedger.frx":236E
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
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "New &Record "
            Key             =   "N"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
      Begin VB.TextBox txtJobCard 
         Height          =   285
         Left            =   9960
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDeptCode 
         Height          =   285
         Left            =   10560
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   735
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   2640
      TabIndex        =   4
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
      TabIndex        =   2
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
Attribute VB_Name = "frmODASMLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsUSED As clsODASMCJobBrief

Private Sub calculateRECEIPTTOTALS()
On Error GoTo err
Exit Sub
err:
        ErrorMessage
End Sub

Private Sub Form_Activate()
    Set rsUSED = New clsODASMCJobBrief
        rsUSED.loadLEDGER
        rsUSED.calculateTOTALINVOICES
        rsUSED.calculateTOTALRECEIPTS
        Set rsUSED = Nothing
        
        showJBITEMS
        showJOBBRIEFRECEIPTs
        showJOBBRIEFINVOICES
        
        disableALLRECORD
    Set rsUSED = Nothing
End Sub
Private Sub enableLISTVIEW()
On Error GoTo err
        
        With frmODASMCosting
            .ListView1.Enabled = True
            .ListView2.Enabled = True
            .ListView3.Enabled = True
            .ListView5.Enabled = True
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub disableLISTVIEW()
On Error GoTo err
        
        With frmODASMCosting
            .ListView1.Enabled = False
            .ListView2.Enabled = False
            .ListView3.Enabled = False
            .ListView5.Enabled = False
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub disableFRAMES()
On Error GoTo err
        
        With frmODASMCosting
            .Frame2.Enabled = False
            .Frame6.Enabled = False
        End With

Exit Sub

err:
    ErrorMessage
End Sub



Private Sub Form_Load()
        OpenConnection
End Sub


Private Sub Label8_Click()

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
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
'
            frmODASMCosting.txtDeptCode.Text = Item.Text
'            Set rsUSED = New clsODASMCJobBrief
'            rsUSED.loadITEM
'            Set rssed = Nothing
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Terminate()
       Set rsUSED = Nothing
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
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

Private Sub ListView5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView5.SortKey = ColumnHeader.Index - 1
    ListView5.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
'Set rsUSED = New clsODASMCJobBrief
   
 With frmODASMLedger
    
    Set rsreceipt = New clsReceipting1
    Set rsReceiptDetails = New clsReceiptDetails1
    
    Select Case Button.Key
    Case "N"
    
                Select Case Button.Caption
                    Case "New &Record "
                        If editRECORD Then Exit Sub
                        disableALLRECORD
                        rsUSED.enableRECORD
                        NewRecord = True: Button.Caption = "&Compute Costs"
                        enableLISTVIEW
                    Case "&Save Record "
                        If NewRecord Then
                                    
                                    rsUSED.updateJobBriefCosts
                                    Button.Caption = "New &Record ": Button.Image = 2
'                                    rsUSED.updateRECORD
'                                    Button.Caption = "NE&XT ITEM"
'                                    .Toolbar1.Buttons(3).Caption = "FINISH"
                        End If
'
'                     Case "NE&XT ITEM"
'                            Button.Caption = "&Save Record ": Button.Image = 4
'                            rsUSED.clearRECORD
'                            rsUSED.enableRECORD
'
'                            NewRecord = False
                    Case Else
                            Exit Sub
                End Select
    
      Case "E"
      Case "S"
               Select Case Button.Caption
                    Case "NEXT ATTACH"
                            Button.Caption = "&Save Record ": Button.Image = 4
                            disableFRAMES
                    Case "&Save Record "
                            rsUSED.updateRECORD
                            disableALLRECORD
                            .Toolbar1.Buttons(3).Caption = "FINISH"
                End Select
                
        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False:
        Case "P"
        Case "F"
            Me.HelpCommonDialog.DialogTitle = "Using the Main System"
            Me.HelpCommonDialog.HelpFile = App.HelpFile
            Me.HelpCommonDialog.HelpContext = 71
            Me.HelpCommonDialog.HelpCommand = cdlHelpContext
            Me.HelpCommonDialog.ShowHelp
        Case Else
            Exit Sub
        End Select

    End With
    
Set rsUSED = Nothing
Set rsreceipt = Nothing
Set rsReceiptDetails = Nothing

Exit Sub
err:
    ErrorMessage

End Sub

Private Sub txtItemPrice_LostFocus()
On Error GoTo err
    
    With frmODASMCosting
        .txtNetItemPrice.Text = CDbl(.txtItemQuantity.Text) * CDbl(.txtItemPrice.Text)
        .txtNetItemPrice.Text = CDbl(.txtNetItemPrice.Text)
    End With

Exit Sub
err:
    ErrorMessage
End Sub


Private Sub UpDownDuration_Change()
On Error GoTo err
        With frmODASMCosting
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
On Error GoTo err
        With frmODASMCosting
            .txtItemQuantity.Text = .UpDownQuantity.Value
        End With
Exit Sub

err:
    ErrorMessage
End Sub
