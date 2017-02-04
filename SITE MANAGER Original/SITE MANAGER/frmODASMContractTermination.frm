VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASMContractTermination 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contract Termination"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtContractNo 
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtPlotNo 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONTRACTS FOR THE SELECTED PLOT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   14055
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   1164
      ButtonWidth     =   3307
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1(1)"
      DisabledImageList=   "ImageList1(1)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New Record"
            Key             =   "N"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
            Object.Visible         =   0   'False
            Caption         =   "&Help System  "
            Key             =   "H"
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
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   0
         Left            =   10920
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
               Picture         =   "frmODASMContractTermination.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":067A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":0ACC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":0DE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":1460
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":1ADA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":1F2C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   1
         Left            =   0
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
               Picture         =   "frmODASMContractTermination.frx":25A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":2C20
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":3162
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":35B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":38CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":3F48
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":45C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMContractTermination.frx":4A14
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Contract No:"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Plot No:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmODASMContractTermination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
checkOne Item, Me.ListView1
If Item.Checked = True Then
    Item.Selected = True
    Me.txtContractNo.Text = Item.Text
Else
    Item.Selected = False
    Me.txtContractNo.Text = ""

End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Selected = True Then
    Item.Checked = True
Else
    Item.Checked = False
End If
checkOne Item, Me.ListView1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "&New Record"
                        baddRECORD = True
                        enableALLRECORD
                        Button.Caption = "&Terminate Contract": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Terminate Contract"
                            ValidateRECORD
                            If bSaveRECORD = True Then
                                SaveRecord
                                baddRECORD = False: Button.Caption = "&New Record": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                        End If
                   End Select
            Case "E"
                Select Case Button.Caption
                    Case "&Edit/Change "
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                
                    Case "Save &Changes "
                    
                      
                        
                        beditRECORD = False: Button.Caption = "&Edit/Change ": Button.Image = 6
                    End Select
                    
            Case "P"
                
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                        bsearchRECORD = True
                        CurrentRecord = InputBox("Enter the Plot No")
                        If CurrentRecord = "" Then Exit Sub
                        Me.txtPlotNo.Text = CurrentRecord
                        strSQL = "SELECT LA.COntractNo,LA.PlotNo,A.COmpanyName,COmmencementdate,ExpiryDate,LeaseDuration,Terminated,TerminationDate FROM ODASMLeaseAgreement LA INNER JOIN ODASPAccount A ON A.AccountNo=LA.AccountNo WHERE PlotNo LIKE '" & CurrentRecord & "'"
                        FillList strSQL, Me.ListView1
                        'Button.Caption = "Delete"
                    Case "Cancel"
                        cancelCMD
                        baddRECORD = False: Button.Caption = "&New Record": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                        
                    Case "Delete"
                          
                        Button.Caption = "&Search/Find "
                     beditRECORD = False: Button.Caption = "Delete": Button.Image = 6
                End Select
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 3
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 6
                    .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                    NewRecord = False: editRECORD = False: 'MyCommonData.ClearTheScreen

            Case "H"
                Me.HelpCommonDialog.DialogTitle = "Using the Main System"
                Me.HelpCommonDialog.HelpFile = App.HelpFile
                Me.HelpCommonDialog.HelpContext = 19
                Me.HelpCommonDialog.HelpCommand = cdlHelpContext
                Me.HelpCommonDialog.ShowHelp
        End Select
    End With
Exit Sub
err:
ErrorMessage
End Sub


Private Sub ValidateRECORD()
bSaveRECORD = False
If Me.ListView1.View = lvwList Then
    Exit Sub
End If

For i = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(i).Checked = True Then
                bSaveRECORD = True
                Exit For
        End If
Next i

If Me.txtContractNo.Text = "" Then
        bSaveRECORD = False
End If

If bSaveRECORD = False Then
    MsgBox "Select at least one record to terminate", vbExclamation
End If
End Sub


Private Sub SaveRecord()
On Error GoTo errMSG
strSQL = "UPDATE ODASMLeaseAgreement SET Terminated='Y',TerminationDate='" & Format(Date, "yyyy/MM/dd") & "',TerminatedBy='" & CurrentUserName & "' WHERE ContractNo LIKE '" & Me.txtContractNo & "'"
Set rsLease = cnCOMMON.Execute(strSQL)
Exit Sub
errMSG:
    ErrorMessage
End Sub
