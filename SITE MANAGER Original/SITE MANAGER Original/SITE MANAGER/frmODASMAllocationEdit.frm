VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMAllocationEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Contract Setup"
   ClientHeight    =   6495
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10605
   Icon            =   "frmODASMAllocationEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   5280
      TabIndex        =   5
      Top             =   720
      Width           =   4935
      Begin VB.TextBox txtContractNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPickerAgreementDate 
         Height          =   315
         Left            =   2760
         TabIndex        =   21
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57278465
         CurrentDate     =   38365
      End
      Begin VB.TextBox txtSignedBy 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txtWitnessCoy 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtWitnessLandLord 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   3015
      End
      Begin VB.CheckBox chkDeallocate 
         Caption         =   "De Allocate?"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPlotNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtAgreementDate 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtNames 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtLandLordNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Contract No"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   150
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Signed By"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2670
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Coy's Witness"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2310
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Land Lord's Witness"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1950
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Plot No"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Agreement Date"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Names"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Land Lord No"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   510
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties Belonging to this LandLord"
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
      Height          =   2655
      Left            =   480
      TabIndex        =   3
      Top             =   3720
      Width           =   9735
      Begin MSComctlLib.ListView ListView2 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4048
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
            Picture         =   "frmODASMAllocationEdit.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocationEdit.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocationEdit.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocationEdit.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocationEdit.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocationEdit.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAllocationEdit.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "All Active Land Lords"
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
      Height          =   3015
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4695
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4683
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10605
      _ExtentX        =   18706
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
         Left            =   9120
         Top             =   0
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
Attribute VB_Name = "frmODASMAllocationEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsALLOCATION As clsODASAllocation

Private Sub Form_Activate()
        Set rsALLOCATION = New clsODASAllocation
        rsALLOCATION.loadRECORD
        Set rsALLOCATION = Nothing
        getLANDLORDS
        showALLLandLORDSites
        disableFRAME
End Sub

Private Sub loadDEFAULTS()
'On Error GoTo err
        With frmODASMAllocationEdit
            .txtAgreementDate.Text = Date
        End With

Exit Sub

err:
    ErrorMessage
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
            
            frmODASMAllocationEdit.txtLandLordNo.Text = Item.Text
            frmODASMAllocationEdit.txtNames.Text = Item.SubItems(1)
            showALLLandLORDSites

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
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
            
            frmODASMAllocationEdit.txtSiteNo.Text = Item.Text
                    Set rsALLOCATION = New clsODASAllocation
                    rsALLOCATION.loadRECORD
                    Set rsALLOCATION = Nothing
            
            Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub disableFRAME()
'On Error GoTo err
    
    With frmODASMAllocationEdit
        .Frame1.Enabled = False
'        .Frame2.Enabled = False
'        .Frame3.Enabled = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub enableFRAME()
'On Error GoTo err
    
    With frmODASMAllocationEdit
        .Frame1.Enabled = False
        .Frame2.Enabled = True
        .Frame3.Enabled = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''On Error GoTo Err
        
        With frmODASMAllocationEdit
        Set rsALLOCATION = New clsODASAllocation

        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                        If EditRecord Then Exit Sub
                        .ListView2.ListItems.Clear:
                        rsALLOCATION.enableRECORD
                        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                        enableFRAME
                    Case "&Save Record "
                        If NewRecord Then
                            rsALLOCATION.updateRECORD

                            .Toolbar1.Buttons(3).Caption = "FINISH"
                            
                        End If
                        
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                Case Else
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
                If Screen.ActiveForm.txtContractNo.Text = Empty Then Exit Sub
                Load frmODASRContract1
                frmODASRContract1.Show 1, Me
        Case "F"
     
     
        Case Else
            Exit Sub
        End Select
        Set rsALLOCATION = Nothing

End With
Exit Sub
err:
    ErrorMessage

End Sub


