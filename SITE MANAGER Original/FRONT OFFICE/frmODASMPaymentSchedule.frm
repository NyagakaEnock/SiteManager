VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmODASMCreditAuthorization 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Authorization"
   ClientHeight    =   6975
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmODASMPaymentSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Current Job Briefs"
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
      Left            =   5880
      TabIndex        =   36
      Top             =   600
      Width           =   5895
      Begin MSComctlLib.ListView ListView6 
         Height          =   2415
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4260
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
   Begin VB.Frame Frame3 
      Caption         =   "Authorization Details"
      Height          =   3495
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   5655
      Begin VB.TextBox txtRemarkRequired 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   48
         Top             =   2520
         Width           =   270
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3360
         TabIndex        =   45
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtCurrentPeriod 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4080
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAuthorizationNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4080
         TabIndex        =   41
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboGuarantorType 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1560
         TabIndex        =   39
         Top             =   945
         Width           =   3855
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2880
         TabIndex        =   38
         Top             =   1710
         Width           =   2535
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtAuthorizationDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtAuthorizedBy 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtGuarantor 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtPercentAuthorized 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Top             =   1710
         Width           =   1095
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1560
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Label17 
         Caption         =   "Remark Required"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Guarantor"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Status"
         Height          =   255
         Left            =   2880
         TabIndex        =   46
         Top             =   2175
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Current Period"
         Height          =   255
         Left            =   2880
         TabIndex        =   44
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Authorization No"
         Height          =   255
         Left            =   2880
         TabIndex        =   42
         Top             =   615
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Left            =   2640
         TabIndex        =   40
         Top             =   1725
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "Account No"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Authorization Date"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Authorized By"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2175
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Guaranteed By"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   975
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "Percent Authorized"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1725
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2895
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5655
      Begin VB.TextBox txtBalance 
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
         Height          =   285
         Left            =   3720
         TabIndex        =   21
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtExpiryDate 
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
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Top             =   1005
         Width           =   1695
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   630
         Width           =   3975
      End
      Begin VB.TextBox txtPriceInclusive 
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
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtJobBriefDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1770
         Width           =   3975
      End
      Begin VB.TextBox txtJobBriefNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCommencementDate 
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
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1005
         Width           =   1335
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1395
         Width           =   3975
      End
      Begin VB.Label Label19 
         Caption         =   "Balance"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   2175
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Total Cost (Incl)"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Expires"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Product"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Price Inclusive"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2175
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Brief Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Job Brief No"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Desc of Order"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1785
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Commencement"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Customer"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1410
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Receipts"
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
      Height          =   3495
      Left            =   5880
      TabIndex        =   1
      Top             =   3360
      Width           =   5895
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5530
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
      Left            =   6840
      Top             =   240
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
            Picture         =   "frmODASMPaymentSchedule.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentSchedule.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentSchedule.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentSchedule.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentSchedule.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentSchedule.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentSchedule.frx":236E
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
         Top             =   -120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
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
Attribute VB_Name = "frmODASMCreditAuthorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCREDIT As clsODASCreditAuthorization

Private Sub cboGuarantorType_GotFocus()
        selectGuarantorGotFocus
End Sub

Private Sub cboGuarantorType_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboGuarantorType_LostFocus()
        selectGuarantorLostFocus
        loadGuarantor
        Set rsCREDIT = New clsODASCreditAuthorization
        rsCREDIT.checkCASHSTATUS
        Set rsCREDIT = Nothing
        
End Sub

Private Sub Form_Activate()
        Set rsCREDIT = New clsODASCreditAuthorization
        disableALLRECORD
        rsCREDIT.loadRECORD
        loadGuarantor
        showBRIEFINACCOUNT
        showBRIEFRECEIPTS
End Sub

Private Sub Form_Initialize()
        Set rsCREDIT = New clsODASCreditAuthorization
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub Form_Terminate()
        Set rsCREDIT = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set rsCREDIT = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
''oN ERROR GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''oN ERROR GoTo err
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
            
            Screen.ActiveForm.txtRequisitionNo.Text = Item.Text
            Set rsCREDIT = New clsODASCreditAuthorization
            rsCREDIT.loadREQUISTION
            Set rsCREDIT = Nothing

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'''oN ERROR GoTo Err
        
        With frmODASMReceiveinvoice
        Set rsCREDIT = New clsODASCreditAuthorization
        
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                            If editRECORD Then Exit Sub
                            NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                            rsCREDIT.enableRECORD
                    Case "&Save Record "
                    
                            bsaveRECORD = False
                            rsCREDIT.updateRECORD
                                    
                            If bsaveRECORD = True Then
                                        bsaveRECORD = False
                                        .Toolbar1.Buttons(2).Caption = "New &Record "
                                        .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                        .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                            End If
                    
                    Case "&NEXT ITEM "
                            
                            .Toolbar1.Buttons(1).Caption = "&Save Record"
                            rsCREDIT.enableRECORD
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                
                Case "&Save Record "

                        bsaveRECORD = False
                        rsCREDIT.validateRECORD
                        
                        If bsaveRECORD = True Then
                                rsCREDIT.updateRECORD
                                If bsaveRECORD = False Then
                                          .Toolbar1.Buttons(2).Caption = "New &Record "
                                          .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                          .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                                End If
                        End If
                
                Case "&NEXT ITEM "
                            .Toolbar1.Buttons(3).Caption = "&Save Record "
                            rsCREDIT.enableRECORD
                            'rsCREDIT.clearRECORD
                Case Else
            End Select
        
        Case "S"
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD

        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD
        Case "P"
        Case "F"
     
     
        Case Else
            Exit Sub
        End Select
        
        Set rsCREDIT = Nothing
        
End With
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub txtPercentAuthorized_LostFocus()
    Set rsCREDIT = New clsODASCreditAuthorization
    rsCREDIT.calculateDEPOSIT
    Set rsCREDIT = Nothing
End Sub
