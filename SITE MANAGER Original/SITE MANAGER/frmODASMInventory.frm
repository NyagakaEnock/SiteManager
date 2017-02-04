VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASMInventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Details"
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmODASMInventory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
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
            Picture         =   "frmODASMInventory.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMInventory.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMInventory.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMInventory.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMInventory.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMInventory.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMInventory.frx":236E
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
   Begin VB.Frame Frame3 
      Height          =   5655
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   10815
      Begin VB.TextBox txtLastLPOStatus 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtLastLPONo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtUnitCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   5280
         TabIndex        =   28
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtStockOnOrder 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPickeUpdate 
         Height          =   315
         Left            =   10320
         TabIndex        =   25
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57475073
         CurrentDate     =   38298
      End
      Begin VB.TextBox txtLastIssueDate 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtReOrderStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtReorderLevel 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtStockIssued 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtStockReserved 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   5280
         TabIndex        =   9
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtLastLPODate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   8640
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtStockReceived 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtLastUpDate 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   8640
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtProductDescription 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtStockOnHand 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4471
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Last LPO Status"
         Height          =   195
         Left            =   7200
         TabIndex        =   33
         Top             =   2490
         Width           =   1155
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Last LPO No"
         Height          =   195
         Left            =   7200
         TabIndex        =   31
         Top             =   2010
         Width           =   915
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Unit Code"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   29
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Stock On Order"
         Height          =   195
         Left            =   3960
         TabIndex        =   27
         Top             =   1050
         Width           =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Last Issue Date"
         Height          =   195
         Left            =   7200
         TabIndex        =   24
         Top             =   1530
         Width           =   1110
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Re-Order Status"
         Height          =   195
         Left            =   3960
         TabIndex        =   23
         Top             =   2010
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Re-Order level"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2010
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Stock Reserved"
         Height          =   195
         Left            =   3960
         TabIndex        =   20
         Top             =   2490
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Last LPO Date"
         Height          =   195
         Left            =   7200
         TabIndex        =   19
         Top             =   1050
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Stock Received"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1530
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Update"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7200
         TabIndex        =   17
         Top             =   570
         Width           =   945
      End
      Begin VB.Label Label13 
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Product Description"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   14
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stock On Hand"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1050
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Stock Issued"
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   1530
         Width           =   930
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
Attribute VB_Name = "frmODASMInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsINVENTORY As New clsODASInventory

Private Sub DTPickerAcquisitionDate_Change()
'On Error GoTo err
    With frmODASMSiteRegistration
            .DTPickerAcquisitionDate.Value = .txtAcquisitionDate.Text
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub DTPickerCommencementDate_Change()
'On Error GoTo err
    With frmODASMSiteRegistration
            .DTPickerCommencementDate.Value = .txtCommencementDate.Text
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        disableALLRECORD
        Set rsINVENTORY = New clsODASInventory
        rsINVENTORY.loadRECORD
        Set rsINVENTORY = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''On Error GoTo Err
        Set rsINVENTORY = New clsODASInventory
        
        With frmODASMSiteRegistration
        
        Select Case Button.Key
        Case "N"
            Select Case Button.Caption
                Case "New &Record "
                    Set rsINVENTORY = New clsODASInventory
                    rsINVENTORY.loadRECORD
                    Set rsINVENTORY = Nothing

                    If EditRecord Then Exit Sub
                    .ListView1.ListItems.Clear: enableALLRECORD
                    NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                    
                Case "&Save Record "
                    If NewRecord Then
                            rsINVENTORY.SaveRECORD
                            .Toolbar1.Buttons(3).Caption = "FINISH"
                    End If
                    
                 Case "NE&XT ITEM"
                      
                Case Else
                    Exit Sub
                End Select
            
        Case "E"
                Select Case Button.Caption
                Case "&Edit/Change "
                    If NewRecord Then Exit Sub
                            If .txtSiteNo.Text = Empty Then
                                MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                               .txtSiteNo.SetFocus
                            Else
                               .txtSiteNo.Locked = True
                                Button.Caption = "Save &Changes ": Button.Image = 4
                                EditRecord = True
                            End If
                Case "Save &Changes "
                    If EditRecord Then
                        If ValidRecord Then
                                rsINVENTORY.SaveRECORD
                                Set rsEditRecord = Nothing
                                .txtSiteNo.Locked = False: EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 5
                        End If
                    End If
                 Case "FINISH"
                     If ValidMainRecord Then
                            .Toolbar1.Buttons(2).Caption = "New &Record "
                            .Toolbar1.Buttons(2).Image = 2
                            .Toolbar1.Buttons(3).Image = 5
                            .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                         
                    End If
                Case Else
                   
                    Exit Sub
                End Select
        Case "S"
        Case "R"
                If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 2
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 5
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
        Case "P"
                Load frmRptAdvertPrintOut
                frmRptAdvertPrintOut.Show 1, frmODASMSiteRegistration
        Case "F"
             
             
        Case Else
            Exit Sub
        End Select
        End With

Set rsINVENTORY = Nothing

        
Exit Sub
err:
    ErrorMessage

End Sub


Private Sub UpDownLeasePeriod_Change()
'On Error GoTo err
        With frmODASMSiteRegistration
            .txtLeaseDuration.Text = .UpDownLeasePeriod.Value
        End With

Exit Sub

err:
    ErrorMessage
End Sub
