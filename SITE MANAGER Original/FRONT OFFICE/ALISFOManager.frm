VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ALISFOManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O.D.A.S. PLUS ~ [FRONT OFFICE MODULE]"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   840
   ClientWidth     =   15705
   Icon            =   "ALISFOManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ALISFOManager.frx":0442
   ScaleHeight     =   8340
   ScaleWidth      =   15705
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALISFOManager.frx":5B80
            Key             =   "C"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALISFOManager.frx":5FD4
            Key             =   "O"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALISFOManager.frx":6426
            Key             =   "F"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7335
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12938
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
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   12938
      _Version        =   393217
      Indentation     =   531
      LineStyle       =   1
      Style           =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
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
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15705
      _ExtentX        =   27702
      _ExtentY        =   1058
      ButtonWidth     =   3254
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Receipts"
            Key             =   "R"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Receive Invoice"
            Key             =   "RI"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prepare Voucher"
            Key             =   "PV"
            ImageIndex      =   42
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prepare Cheque"
            Key             =   "C"
            ImageIndex      =   41
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Authorize Credit"
            Key             =   "AC"
            ImageIndex      =   44
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prepare Invoice"
            Key             =   "PI"
            ImageIndex      =   43
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User Guide/Help"
            Key             =   "HLP"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10560
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtcurrentPeriod 
         Height          =   285
         Left            =   15840
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtTaskDetails 
         Height          =   285
         Left            =   14040
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9000
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   45
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":6878
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":6B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":720C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":7886
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":7F00
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":857A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":8BF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":926E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":98E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":9F62
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":A5DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":AC56
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":B2D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":B94A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":BFC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":C2DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":C5F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":CA4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":CE9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":D2EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":D740
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":DB92
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":DFE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":E436
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":E888
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":ECDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":F12C
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":F57E
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":F9D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":FE2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":1027C
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":106CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":10B20
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":10F72
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":113C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":11816
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":11C68
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":120BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":123D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":12B26
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":12F78
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":133CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":13820
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":13C74
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISFOManager.frx":140C8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtTASK 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   12600
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtCompanyCode 
         Height          =   285
         Left            =   13200
         TabIndex        =   6
         Text            =   "MAGNATE"
         Top             =   120
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   7950
      Width           =   15705
      _ExtentX        =   27702
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   5998
            MinWidth        =   5998
            Picture         =   "ALISFOManager.frx":1451C
            Text            =   "OutDoor Adverizing System"
            TextSave        =   "OutDoor Adverizing System"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7057
            MinWidth        =   7057
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7057
            MinWidth        =   7057
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   7
            Alignment       =   2
            Enabled         =   0   'False
            Object.Width           =   7057
            MinWidth        =   7057
            TextSave        =   "KANA"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileEndSession 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuasdsad 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActionClearData 
         Caption         =   "&Clear Datasheet"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionPrintSelect 
         Caption         =   "Print &Selected Records"
      End
      Begin VB.Menu mnui8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionPrintAll 
         Caption         =   "Print &All Displayed"
      End
      Begin VB.Menu mnuwerewr 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuStatements 
      Caption         =   "Customer Statements"
      Begin VB.Menu mnuOpening 
         Caption         =   "Opening Balances"
      End
      Begin VB.Menu mnurewfdxzfdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Statement"
      End
      Begin VB.Menu hghghgh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBatchUpdate 
         Caption         =   "Batch Rent Update"
      End
      Begin VB.Menu mnulinooo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClientBills 
         Caption         =   "Client Billing Listing"
      End
   End
   Begin VB.Menu mnuHelpSystem 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Help &Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Search for &Help On..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A&bout the System..."
      End
   End
End
Attribute VB_Name = "ALISFOManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SysManage As clsSysManager
Private rsREQUISITION As clsALISCheque
Private rsPAYMENT As clsALISPaymentRequisition
Private rsreceipt As clsALISReceipt
Dim rsOPERATION As clsODASOperation: Private rsClaimApproval As clsALISApproval

Private Sub ListView1_BeforeLabelEdit(cancel As Integer)
  cancel = 1
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_DblClick()
On Error GoTo err
    
            With ALISFOManager
                    
                CurrentRecord = Trim(Me.ListView1.SelectedItem.Text)
                globalDepartmentCode = Empty
                bapproveRECORD = False
                bCostingsApproval = False
                bCostingsAuthorization = False
                bApproveVOUCHER = False
                bauthorizeVOUCHER = False
                bApproveCheque = False
                authorizecheque = False
                bapproveINVOICE = False
                bAuthorizationinvoice = False
                                
                frmODASMOperation.txtApplicationNo.Text = CurrentRecord
                globalDepartmentCode = Screen.ActiveForm.ListView1.SelectedItem.SubItems(1)
                    
                    Set rsOPERATION = New clsODASOperation

                    Select Case (.txtTASK)
                        
                       Case "PC8"
                                    bCostingsApproval = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bCostingsApproval = False Then Exit Sub
                        Case "PC9"
                                    bCostingsAuthorization = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bCostingsAuthorization = False Then Exit Sub
                        Case "R41": bApproveVOUCHER = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bApproveVOUCHER = False Then Exit Sub
                        
                        Case "R42": bauthorizeVOUCHER = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bauthorizeVOUCHER = False Then Exit Sub
                                    
                        Case "K41": bApproveCheque = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bApproveCheque = False Then Exit Sub
                                    
                        Case "K42": bAuthorizeCheque = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bAuthorizeCheque = False Then Exit Sub
                                    
                        Case "I5": bapproveINVOICE = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bapproveINVOICE = False Then Exit Sub
                        Case "T5": bapproveINVOICE = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bapproveINVOICE = False Then Exit Sub
                                    
                        Case "I6": bauthorizeINVOICE = True
                                    rsOPERATION.checkAPPROVEDDISCHARGE
                                    If bauthorizeINVOICE = False Then Exit Sub
                                    
                    
                    End Select
                    If bapproveRECORD = True Then rsOPERATION.approveOPERATION
                    Set rsOPERATION = Nothing

            End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView2_DblClick()
On Error GoTo err
        With ALISFOManager

                initBooleans
                CurrentRecord = Trim(Me.ListView1.SelectedItem.Text)
                
                frmODASMOperation.txtApplicationNo.Text = CurrentRecord
                
                Set rsClaimApproval = New clsALISApproval
    
                Select Case (.txtTASK.Text)
                
                Case "C5"
                        Load frmALISMLedgerDetails
                        frmALISMLedgerDetails.Show 1, ALISFOManager
                Case "C9"
                        Load frmALISMLedgerDetails
                        frmALISMLedgerDetails.Show 1, ALISFOManager
                        strSQL = "Select ALISMReceiptNew.txtReceiptNo, ALISMReceiptNew.Receiptdate, ALISMReceiptNew.ReceiptAmount, ALISMReceiptNew.AccountingPeriod, ALISMReceiptNew.Payer, ALISMReceiptNew.PaymentMethod, ALISMReceiptNew.ChequeNo, ALISMReceiptNew.BankNo from ALISMReceiptNew, ALISPDefaults Where ALISMReceiptNew.AccountingPeriod = ALISPDefaults.CurrentPeriod order by receiptno;"
                        rsreceipt.getRECEIPT
    
                Case "R7":
                        bApproveVOUCHER = True
                        rsClaimApproval.checkAPPROVEDDISCHARGE
                        If bapproveRECORD = False Then Exit Sub
                        rsClaimApproval.approveCLAIM
    
                Case "R8":
                        bauthorizeVOUCHER = True
                        rsClaimApproval.checkAPPROVEDDISCHARGE
                        If bauthorizeVOUCHER = False Then Exit Sub
                        rsClaimApproval.approveCLAIM
                Case "K7":
                        bApproveCheque = True
                        rsClaimApproval.checkAPPROVEDDISCHARGE
                        If bApproveCheque = False Then Exit Sub
                        rsClaimApproval.approveCLAIM
    
                Case "K8":
                        bAuthorizeCheque = True
                        rsClaimApproval.checkAPPROVEDDISCHARGE
                        If bAuthorizeCheque = False Then Exit Sub
                        rsClaimApproval.approveCLAIM
                
                End Select
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
   
        initBooleans
        
        With ALISFOManager
        
            If Item.Checked = True Then
                
                j = Screen.ActiveForm.ListView1.ListItems.Count
                
                If j = 0 Then Exit Sub
                
                For i = 1 To j
                    If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                       Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                    End If
                Next i

                Select Case (.txtTASK.Text)
                    Case "INST"
                        CurrentRecord = Item
                        Load frmODASRPaymentInstallments
                        frmODASRPaymentInstallments.Show vbModal
                    Case "D6"
                        Load frmODASPPaymentMethod
                        frmODASPPaymentMethod.txtPaymentMethod.Text = Item.Text
                        frmODASPPaymentMethod.Show 1, Me
                    
                    Case "D10"
                        Load frmODASPDuration
                        frmODASPDuration.txtDurationMode.Text = Item.Text
                        frmODASPDuration.Show 1, Me
    
                    Case "D8"
                        Load frmODASPGuarantor
                        frmODASPGuarantor.txtGuarantorType.Text = Item.Text
                        frmODASPGuarantor.Show 1, Me

                    Case "D7"
                        Load frmODASPLandRate
                        frmODASPLandRate.txtTownCode.Text = Item.Text
                        frmODASPLandRate.txtTown.Text = Item.SubItems(1)
                        frmODASPLandRate.Show 1, Me
                        
                    Case "R2":
                        frmODASMReceiveinvoice.txtLPONo.Text = Item.Text
                        frmODASMReceiveinvoice.txtRequisitionNo.Text = Item.SubItems(2)
                        
                        Load frmODASMReceiveinvoice
                        frmODASMReceiveinvoice.Show 1, Me
                    Case "R7"
                        CurrentRecord = Item
                        Load frmPayRequisition
                        frmPayRequisition.Show 1, Me

                    Case "C1"
                        Load frmODASMReceipt
                        frmODASMReceipt.txtPaymentMethod.Text = Item
                        CurrentRecord = Item.Text
                        frmODASMReceipt.Show 1, ALISFOManager

                    Case "T4"
                        
                        Load frmODASMAccounts
                        frmODASMAccounts.txtJobBriefNo.Text = Item.SubItems(1)
                        frmODASMAccounts.txtInvoiceReference.Text = Item.Text
                        frmODASMAccounts.txtInstallmentNo.Text = Item.SubItems(7)
                        frmODASMAccounts.Show 1, Me
                    Case "STT"
                        CurrentRecord1 = Item.SubItems(1)
                        frmRCustomerStatement.Show vbModal

                    Case "T7"
                        Load frmODASmInvoiceIssuance
                        frmODASmInvoiceIssuance.txtInvoiceNo.Text = Item.Text
                        frmODASmInvoiceIssuance.Show 1, Me
       
                    Case "D5"
                        Load frmALISPReversalType
                        frmALISPReversalType.txtReversalType = Item
                        frmALISPReversalType.Show 1, Me
                
                    Case "C6"
                        frmALISMSuspense.cboDocumentNo.Text = Item.Text
                        frmALISMSuspense.txtTransactionAmount.Text = Item.SubItems(2)
                        Load frmALISMSuspense
                        frmALISMSuspense.Show 1, ALISFOManager
                    Case "C9"
                        frmALISMReceiptCopy.txtReceiptNo.Text = Item.Text
                        Load frmALISMReceiptCopy
                        frmALISMReceiptCopy.Show 1, ALISFOManager
       
                    Case "T1"
                        frmODASMCreditAuthorization.txtJobBriefNo.Text = Item.Text
                        Load frmODASMCreditAuthorization
                        frmODASMCreditAuthorization.Show 1, ALISFOManager
                
                    Case "T3"
                        frmODASMReceiptSchedule.txtJobBriefNo.Text = Item.Text
                        Load frmODASMReceiptSchedule
                        frmODASMReceiptSchedule.Show 1, ALISFOManager
                    Case "T11"
                        Load frmODASMPaySchedule
                        frmODASMPaySchedule.txtJobBriefNo.Text = Item.Text
                        frmODASMPaySchedule.Show vbModal
                    Case "C5"
                        
                    Case "R1":
                        .txtTASK = "R1"
                        frmODASMVoucher.cboPaymentCode.Text = Item.Text
                        Load frmODASMVoucher
                        frmODASMVoucher.Show 1, Me
                
                    Case "R7":
                        .txtTASK = "R7"
                        bapproveREQUISITION = True
                        
                        If .ListView1.Checkboxes = True Then
                                frmODASMVoucher.txtVoucherNo.Text = Item.Text
                                Load frmODASMVoucher
                                frmODASMVoucher.Show 1, Me
                        End If
                
                    Case "R8":
                        .txtTASK = "R8"
                        
                        bAuthorizeREQUISITION = True
                        
                        frmODASMVoucher.txtVoucherNo.Text = Item.Text
                        Load frmODASMVoucher
                        frmODASMVoucher.Show 1, Me
                        
                   Case "R42"
                        CurrentRecord = Item
                        Load frmPayRequisition
                        frmPayRequisition.Show 1, Me

                    Case "K2":
                        .txtTASK = "K2"
                        bmakePAYMENT = True
                        
                        frmODASMCheck.txtVoucherNo.Text = Item.Text
                        Load frmODASMCheck
                        frmODASMCheck.Show 1, Me
                
                    Case "K7":
                        .txtTASK = "K7"
                        
                        bApproveCheque = True
                        frmODASMCheck.txtChequeNo.Text = Item.Text
                        Load frmODASMCheck
                        frmODASMCheck.Show 1, Me
            
                    Case "K8":
                        .txtTASK = "K8"
                        bAuthorizeCheque = True
                        frmODASMCheck.txtChequeNo.Text = Item.Text
                        Load frmODASMCheck
                        frmODASMCheck.Show 1, Me

                    Case "K6":
                        .txtTASK = "K6"
                        breversePAYMENT = True
                        
                        frmODASMReversePayment.txtContractNo.Text = Item.Text
                        Load frmODASMReversePayment
                        frmODASMReversePayment.Show 1, Me
                
                    Case "K3":
                        .txtTASK = "K3"
                        
                        bissueCHECKS = True
                        frmODASMCheckIssuance.txtChequeNo.Text = Item.Text
                        Load frmODASMCheckIssuance
                        frmODASMCheckIssuance.Show 1, Me
                               
                    Case "C7":
                        .txtTASK = "C7"
                        breversePAYMENT = True
                        frmALISMReverseReceipt.txtReceiptNo.Text = Item.Text
                        Load Screen.ActiveForm
                        frmALISMReverseReceipt.Show 1, Me
                
                    Case "C8":
                        .txtTASK = "C8"
                        breversePAYMENT = True
                        frmALISMReverseReceipt.txtReceiptNo.Text = Item.Text
                        Load frmALISMReverseReceipt
                        frmALISMReverseReceipt.Show 1, Me
                
                    Case "C12"
                        .txtTASK = "C12"
                        frmALISMCorrectReceipt.txtReceiptNo.Text = Item.Text
                        Load frmALISMCorrectReceipt
                        frmALISMCorrectReceipt.Show 1, Me
                    Case "A1"
                        CurrentRecord = Item
                        'frmRCustomerStatement.txtQuotationNo = Item
                        Load frmRCustomerStatement
                        frmRCustomerStatement.Show 1, Me
                End Select

            End If
        End With
              
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub initBooleans()
On Error GoTo err
        bschedulePAYMENT = False
        breversePAYMENT = False
        bissueCHECKS = False
        breversePAYMENT = False
        bApproveCheque = False
        bAuthorizeCheque = False
        bapproveREQUISITION = False
        bAuthorizeREQUISITION = False
        bmakePAYMENT = False
        bapproveRECORD = False
        bscheduledCHEQUES = False
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
    Me.ProgressBar1.Visible = False
    CurrentPeriod
    'ALISFOManager.txtCompanyCode.Text = CompanyCode
    ALISFOManager.txtCurrentPeriod.Text = CurrentPeriod
End Sub

Private Sub Form_Load()
    LoadDEFAULT
    Set SysManage = New clsSysManager
    SysManage.NewTreeSetup
End Sub

Private Sub Form_Resize()
    SysManage.ResizeControls
End Sub

Private Sub Form_Unload(cancel As Integer)
On Error GoTo err
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "UPDATE AdminUserLog SET LogoutDate='" & Format(Date, "MMMM dd,yyyy") & "',LogoutTime='" & FormatDateTime(Now, vbLongTime) & "' WHERE LoginID=" & CLng(CLoginID) & ";", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsEditRecord = Nothing
    cancel = 0
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err
    If Me.ListView1.ListItems.Count = 0 Then Exit Sub

    If Button = 2 Then
        PopupMenu mnuActions, , , , mnuActionClearData
    End If
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuaccperiods_Click()
On Error GoTo err
accperiods = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuActionClearData_Click()
On Error GoTo err
    With Me
    If .ListView1.ListItems.Count = 0 Then Exit Sub
        
        If MsgBox("Clear All the Currently Displayed Records?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Datasheet") = vbNo Then Exit Sub
        .ListView1.ListItems.Clear
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuagents_Click()

End Sub

Private Sub mnuBatchUpdate_Click()
  
  If MsgBox("Carrying Out This Process Wiil Affect A Number Of Records. Do You Want To Continue?", vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!! Batch Files Updating", "ALISELP.HLP", 246) = vbNo Then
         Exit Sub
  End If
     DateIssued = InputBox("Enter the Cut-Off Date for Rent Already Confirmed Cleared.Rent Due Upto This Date Will Be Paid ", "Batch Payment File DD/MM/YYYY")
        If Len(DateIssued) = 0 Or DateIssued = Empty Then
        MsgBox "Either the the date was not entered or the operation was cancelled", vbCritical + vbOKOnly
     Exit Sub
  
        Else
          UdateBatchRent
    End If
End Sub

Private Sub mnuClientBills_Click()
Load frmRptBilling
frmRptBilling.Show vbModal
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

Private Sub mnuFileEndSession_Click()
On Error GoTo err
    If MsgBox("Are you sure you want to end your Current Session?", vbQuestion + vbYesNo, "End Session") = vbYes Then
        Call UpdateLogoutRecord
        ALISFOManager.Hide
            
        Load frmLogin
        frmLogin.Show 1
        
    Else
        Exit Sub
    End If
Exit Sub
err:
If err.Number = 400 Then Resume Next
    ErrorMessage

End Sub

Private Sub mnuHelpAbout_Click()
    SysManage.HelpAbout
End Sub

Private Sub mnuHelpContents_Click()
    SysManage.HelpContents
End Sub

Private Sub mnuHelpIndex_Click()
    SysManage.HelpIndex
End Sub

Private Sub mnuOpening_Click()
    frmODASPCustomerOpeningBal.Show vbModal
End Sub

Private Sub mnuPrint_Click()
'    CurrentRecord2 = InputBox("Please enter the year...", "Accounting year request...")
'    CurrentRecord1 = InputBox("Please enter the Customer name...", "Customer revenue searches...")
'    If Len(CurrentRecord2) = 0 Then Exit Sub
'    If Len(CurrentRecord1) = 0 Then
'        Me.txtTASK = "STT"
        ShowAllClients
'    Else
'        frmRCustomerStatement.Show vbModal
'    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With ALISFOManager
        Select Case Button.Key
        Case "R"
            .txtTASK.Text = "C1"
            showALLPaymentMethods
        Case "RI"
            .txtTASK.Text = "R2"
            showALLLPOS
        Case "PV"
            .txtTASK.Text = "R1"
            GetPaymentCode
        Case "C"
            .txtTASK.Text = "K2"
            GetVoucherAUTHORIZED
        Case "AC"
            .txtTASK.Text = "T1"
            showBRIEFSNOTAUTHORIZED
        Case "PI"
            .txtTASK.Text = "T4"
            showRECEIPTSCHEDULE
        Case "HLP"
            .HelpCommonDialog.DialogTitle = "Using the Main System"
            .HelpCommonDialog.HelpFile = App.HelpFile
            .HelpCommonDialog.HelpContext = 35
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

Private Sub TreeView2_Collapse(ByVal Node As MSComctlLib.Node)
    ALISFOManager.ListView1.ListItems.Clear
    Node.Image = "C"
End Sub

Private Sub TreeView2_Expand(ByVal Node As MSComctlLib.Node)
    ALISFOManager.ListView1.ListItems.Clear
    Node.Image = "O"
End Sub

Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo err
With ALISFOManager

    Select Case Node.Key
    
    Case "A1"
        .txtTASK.Text = "A1"
        ShowAllClients
    
    Case "D10"
        .txtTASK.Text = "D10"
        showDURATIONMODE

    Case "D6"
        .txtTASK.Text = "D6"
        showPAYMENTMETHOD

    Case "T4"
        .txtTASK.Text = "T4"
        showRECEIPTSCHEDULE
        
    Case "T5"
        .txtTASK.Text = "T5"
        showINVOICESprepared
        
    Case "T51"
        .txtTASK.Text = "T51"
        showINVOICESprepared
        
   Case "T52"
        .txtTASK.Text = "T52"
        showINVOICESApproved
        
  Case "I5": bapproveINVOICE = True
        rsOPERATION.checkAPPROVEDDISCHARGE
        If bapproveINVOICE = False Then Exit Sub
                                    
  Case "I6": bauthorizeINVOICE = True
        rsOPERATION.checkAPPROVEDDISCHARGE
        If bauthorizeINVOICE = False Then Exit Sub
                    
    Case "T6"
        frmODASRInvoiceListing.Show vbModal

    Case "T7"
        .txtTASK.Text = "T7"
        showINVOICESAuthorized

    Case "T8"
        Load frmODASRInvoiceListing
        frmODASRInvoiceListing.Show vbModal
    
    Case "T9"
        .txtTASK.Text = "T9"
        showINVOICESPaid
    Case "T10"
         Me.txtTASK.Text = "INST"
         'showALLCLOSEDCOSTINGBRIEFS
         showBRIEFSInstalments
    Case "D7"
        .txtTASK.Text = "D7"
        showALLTOWNS
     
     Case "D8"
        .txtTASK.Text = "D8"
        showALLGUARANTOR

     Case "C1"
        .txtTASK.Text = "C1"
        showALLPaymentMethods
    
    Case "T1"
        .txtTASK.Text = "T1"
        showBRIEFSNOTAUTHORIZED
    
    Case "T3"
        .txtTASK.Text = "T3"
        'showBRIEFSNOTINVOICED
        showCONTRACTNOTINVOICED
     Case "T11"
        .txtTASK.Text = "T11"
        showBRIEFSINVOICED
    Case "C2"
        .txtTASK.Text = "C2"
        showALLRECEIPTS

    Case "C3"
        .txtTASK.Text = "C3"
      
    Case "C4"
        .txtTASK.Text = "C4"
        Load frmALISMReceiptReport
        frmALISMReceiptReport.Show 1, ALISFOManager
        
    Case "C5"
        .txtTASK.Text = "C5"
        getJobBrief
    
    Case "C6"
        .txtTASK.Text = "C6"
        Set rsreceipt = New clsALISReceipt
        rsreceipt.getSUSPENSEPOLICY
        Set rsreceipt = Nothing
    
    Case "T2"
        .txtTASK.Text = "T2"
        showBRIEFSAUTHORIZED

    Case "C8"
        .txtTASK.Text = "C8"
        Set rsreceipt = New clsALISReceipt
        breverseRECEIPT = True
        strSQL = "Select ReceiptNo, Receiptdate, ReceiptAmount, AccountingPeriod, Payer, PaymentMethod, ChequeNo, BankNo from ALISMReceiptNew Where PaymentStatus <> 'PAID';"
        rsreceipt.getRECEIPT
        Set rsreceipt = Nothing
    
    Case "C9"
        .txtTASK.Text = "C9"
        Set rsreceipt = New clsALISReceipt
        breverseRECEIPT = True
        rsreceipt.getRECEIPT
        Set rsreceipt = Nothing
        
    Case "C11"
        .txtTASK.Text = "C11"
        GetSchedule

    Case "R1"
        .txtTASK.Text = "R1"
        GetPaymentCode
    
    Case "R2"
        .txtTASK.Text = "R2"
        showALLLPOS
    
    Case "R3"
        .txtTASK.Text = "R3"
        showALLINVOICESRECEIVED
    Case "10"
        .txtTASK.Text = "R41"
        GetVoucherPrepared
    Case "R11"
        .txtTASK.Text = "R42"
        GetVoucherAPPROVED
    Case "K41"
        .txtTASK.Text = "K41"
        bApproveCheque = True
        GetApprovedChecks
    Case "K42"
        .txtTASK.Text = "K42"
        bAuthorizeCheque = True
        GetAuthorizedChecks
    Case "R7"
        .txtTASK.Text = "R7"
        GetVoucherPrepared
        
    Case "R8"
        .txtTASK.Text = "R8"
        Load frmPayRequisitionListing
        frmPayRequisitionListing.Show 1, Me
    
    Case "R9"
        .txtTASK.Text = "R8"
        GetVoucherAUTHORIZED
 
    Case "K2"
        .txtTASK.Text = "K2"
        GetVoucherAUTHORIZED
        
    Case "K7"
        .txtTASK.Text = "K7"
        bApproveCheque = True
        GetApprovedChecks
    
    Case "K8"
        .txtTASK.Text = "K8"
        
        Set rsREQUISITION = New clsALISCheque
        bAuthorizeCheque = True
        GetAuthorizedChecks

    Case "K3"
        .txtTASK.Text = "K3"
        bissueCHECKS = True
        GetIssuedChecks
    
    Case "K6"
        .txtTASK.Text = "K6"
        getALLContracts

    Case "P1"
        .txtTASK.Text = "P1"
        Load frmOfficeSettings
        frmOfficeSettings.Show 1, ALISFOManager
    
    Case "P3"
        .txtTASK.Text = "P3"
        FindCurrencies
    
    Case "P4"
        .txtTASK.Text = "p4"
        FindVATRate
    Case "D9"
        .txtTASK.Text = "D9"
        FindAccountingPeriods
          
    Case Else
        Exit Sub
    End Select
    
    .txtTaskDetails.Text = Node
End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub UdateBatchRent()
With Me
        Set rsSAVE = New ADODB.Recordset
        strSQL = "Select * from ODASPPlotmast Where RentdueDate>='" & Format(DateIssued, "dd/mm/yyyy") & "'"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
        
        Do While Not rsSAVE.EOF
                
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "Select * From ODASMInstallment where ContractNo = '" & rsSAVE!ContractNo & "' "
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                Else
                        rsCONTROL!PaymentDue = 0
                        rsCONTROL!PaymentFlag = "Y"
                        
                        rsCONTROL.Update
                End If
                
                rsSAVE.MoveNext
        Loop
        
        Set rsSAVE = Nothing
        
End With
End Sub

Public Sub GetVoucherAPPROVED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Voucher No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Voucher Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Account No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMVoucher Where Approved = 'Y' and Authorized = 'N';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!VoucherDate) Then
                                MyList.SubItems(1) = CStr(rsLIST!VoucherDate)
                        End If

                        If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
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
Public Sub GetVoucherPrepared()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Voucher No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Voucher Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Account No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMVoucher Where Prepared = 'Y' and Approved = 'Y' and Authorized='Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!VoucherDate) Then
                                MyList.SubItems(1) = CStr(rsLIST!VoucherDate)
                        End If

                        If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
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

