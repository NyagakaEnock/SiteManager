VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASPCustomerOpeningBal 
   Caption         =   "OPENING BALANCES"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10785
   Icon            =   "frmODASCustomerOpeningBal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmODASCustomerOpeningBal.frx":0442
   ScaleHeight     =   4830
   ScaleWidth      =   10785
   Begin VB.Frame Frame12 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10695
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   10455
         Begin VB.TextBox txtBalance 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   4560
            TabIndex        =   12
            Top             =   600
            Width           =   5775
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2760
            TabIndex        =   10
            Top             =   600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22609921
            CurrentDate     =   38566
         End
         Begin VB.TextBox txtDate 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   840
            TabIndex        =   8
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtCustomerName 
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
            Left            =   3000
            TabIndex        =   5
            Top             =   240
            Width           =   7335
         End
         Begin VB.TextBox txtCustomerNo 
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
            Left            =   840
            TabIndex        =   4
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Opening Balance"
            Height          =   255
            Left            =   3120
            TabIndex        =   11
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Customer"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Customer Begining Balances History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   10455
         Begin MSComctlLib.ListView ListView1 
            Height          =   2415
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   4260
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
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   1164
      ButtonWidth     =   3016
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
            Caption         =   "New &Record"
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
               Picture         =   "frmODASCustomerOpeningBal.frx":0784
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":0DFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":1250
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":156A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":1BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":225E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":26B0
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
               Picture         =   "frmODASCustomerOpeningBal.frx":2D2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":33A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":38E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":3D38
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":4052
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":46CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":4D46
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASCustomerOpeningBal.frx":5198
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmODASPCustomerOpeningBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddNew_Click()
End Sub

Private Sub cmdCancel_Click()
        clearALLRECORD
        disableALLRECORD
        baddRECORD = False
End Sub

Private Sub cmdEdit_Click()
        EditMyRecord
End Sub

Private Sub validateRECORD()
On Error GoTo err
    With Me
    
            bsaveRECORD = False
            
            If .txtCustomerNo.Text = Empty Then
                    MsgBox "The customer number and name is required..."
                .txtCustomerNo.SetFocus
            
            ElseIf .txtDate.Text = Empty Then
                    MsgBox "The opening balance date is reqired..."
                .txtDate.SetFocus
            ElseIf .txtBalance = Empty Then
                    MsgBox "The opening balance must be entered"
                .txtBalance.SetFocus
            Else
                    bsaveRECORD = True
            End If

    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRecord()
On Error GoTo err
    With Me
        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "Select * from ODASMCustomerStatement Where AccountNo = '" & .txtCustomerNo.Text & "' and TransactionDate = '" & Format(.txtDate, "MMMM dd,YYYY") & "' and Reference = 'BAL C/F'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF And rsCONTROL.BOF Then
            rsCONTROL.AddNew
            rsCONTROL!AccountNo = .txtCustomerNo
            rsCONTROL!TransactionDate = Format(.txtDate, "MMMM dd,YYYY")
            rsCONTROL!Reference = "BAL C/F"
        End If
            rsCONTROL!DebitAmount = 0
            rsCONTROL!CreditAmount = .txtBalance
            rsCONTROL!Balance = rsCONTROL!DebitAmount - rsCONTROL!CreditAmount
            rsCONTROL!Details = "Opening Balance"
            rsCONTROL.Update
   End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsCONTROL.CancelUpdate
            rsCONTROL.Requery
    Else
        UpdateErrorMessage
    End If

End Sub


Private Sub cmdUpdate_Click()

        
End Sub

Private Sub cmdSearch_Click()
        
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub DTPicker1_CloseUp()
    Me.txtDate = Me.DTPicker1.Value
End Sub

Private Sub Form_Activate()
    disableALLRECORD
End Sub

Private Sub lblRelationshipCode_Click()

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
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            Me.txtCustomerNo.Text = Item.Text
            Me.txtCustomerName = Item.SubItems(1)
            Me.txtDate = Item.SubItems(2)
            Me.txtBalance = Item.SubItems(3)
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "New &Record"
                        baddRECORD = True
                        clearALLRECORD
                        enableALLRECORD
                        Button.Caption = "&Save Record": Button.Image = 5 ': .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Save Record"
                        bsaveRECORD = False
                        validateRECORD
                        If bsaveRECORD = True Then
                            saveRecord
                                If bsaveRECORD = False Then
                                    disableALLRECORD
                                    baddRECORD = False
                                End If
                        
                        OpeningBalances
                        baddRECORD = False: Button.Caption = "New &Record": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                        End If
                   End Select
            Case "E"
                Select Case Button.Caption
                    Case "&Edit/Change "
                         If baddRECORD Then Exit Sub
                                If .txtCustomerNo.Text = Empty Then
                                MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                               .txtCustomerNo.SetFocus
                                Else
                                beditRECORD = True
                                enableALLRECORD
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                End If
                    Case "Save &Changes "
                        saveRecord
                        OpeningBalances
                        beditRECORD = False: Button.Caption = "&Edit/Change ": Button.Image = 6
                    End Select
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                        SeachCustomer
                End Select
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 3
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 6
                    .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                    NewRecord = False: editRECORD = False ': MyCommonData.ClearTheScreen

            Case "H"
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 19
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
        End Select
    End With
Exit Sub
err:
ErrorMessage

End Sub

Private Sub SeachCustomer()
On Error GoTo err
    With Me
        CurrentRecord = InputBox("Enter the Customer Name to search...")
        If Len(CurrentRecord) = 0 Then Exit Sub
        
            Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open "Select * From ODASPAccount Where CompanyName like '%" & CurrentRecord & "%'", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsFindRecord.EOF And rsFindRecord.BOF Then
                    MsgBox "Sorry! Search completed...System could not find matching a customer ", vbInformation + vbOKOnly, "Missing records"
                Else
                    .txtCustomerName = rsFindRecord!CompanyName
                    .txtCustomerNo = rsFindRecord!AccountNo
                    OpeningBalances
                End If
    End With
Exit Sub
err:
ErrorMessage
End Sub
