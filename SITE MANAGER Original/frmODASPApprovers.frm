VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPApprovers 
   Caption         =   "Approvers"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11340
   Icon            =   "frmODASPApprovers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11340
   Begin VB.Frame Frame12 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11295
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8640
         TabIndex        =   25
         Top             =   2520
         Width           =   1095
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPApprovers.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPApprovers.frx":040C
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3855
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   5415
         Begin MSComctlLib.ListView ListView1 
            Height          =   3495
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   6165
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
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   11055
         Begin VB.TextBox cboOperationType 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   1200
            TabIndex        =   28
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtOperationDescription 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   4560
            TabIndex        =   20
            Top             =   240
            Width           =   6375
         End
         Begin VB.Label Label5 
            Caption         =   "Operation "
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   330
            Width           =   1215
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Properties"
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   11055
         Begin VB.TextBox txtReenterPassword 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   9120
            PasswordChar    =   "x"
            TabIndex        =   4
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtLimitAmount 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   9
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtDateRetired 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   5520
            TabIndex        =   8
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   5520
            PasswordChar    =   "x"
            TabIndex        =   2
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtDateAssigned 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   9120
            TabIndex        =   7
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtUserCode 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtNames 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   4560
            TabIndex        =   5
            Top             =   240
            Width           =   6375
         End
         Begin VB.ComboBox cboStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   6
            Top             =   1215
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "Re-Enter Password"
            Height          =   195
            Left            =   7320
            TabIndex        =   23
            Top             =   810
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Limit"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Date Retired"
            Height          =   195
            Left            =   4560
            TabIndex        =   17
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Left            =   4560
            TabIndex        =   16
            Top             =   810
            Width           =   855
         End
         Begin VB.Label lblBenefitCode 
            Caption         =   "User Code"
            Height          =   315
            Left            =   240
            TabIndex        =   15
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Name"
            Height          =   195
            Left            =   3960
            TabIndex        =   14
            Top             =   330
            Width           =   855
         End
         Begin VB.Label lblStatus 
            Caption         =   "Status"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   1290
            Width           =   735
         End
         Begin VB.Label lblDateAssigned 
            Caption         =   "Date Assigned"
            Height          =   195
            Left            =   7320
            TabIndex        =   12
            Top             =   1260
            Width           =   1215
         End
      End
      Begin VB.Frame Frame14 
         Height          =   3855
         Left            =   5640
         TabIndex        =   10
         Top             =   2640
         Width           =   5535
         Begin MSComctlLib.ListView ListView2 
            Height          =   3495
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   6165
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
            Picture         =   "frmODASPApprovers.frx":050E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPApprovers.frx":0B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPApprovers.frx":0FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPApprovers.frx":12F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPApprovers.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPApprovers.frx":1FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASPApprovers.frx":243A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
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
End
Attribute VB_Name = "frmODASPApprovers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAPPROVERS As clsODASApprovers
Public rsLOAN As ADODB.Recordset

Private Sub cboStatus_gotFocus()
        Set rsAPPROVERS = New clsODASApprovers
        rsAPPROVERS.statusGOTFOCUS
        Set rsAPPROVERS = Nothing
End Sub
Private Sub cbostatus_keypress(Index As Integer)
On Error GoTo Err
        KeyAscii = 0
Exit Sub

Err:
    ErrorMessage

End Sub
Private Sub cboStatus_LostFocus()
On Error GoTo Err

        Set rsAPPROVERS = New clsODASApprovers
        rsAPPROVERS.statusLOSTFOCUS
        Set rsAPPROVERS = Nothing

Exit Sub

Err:
    ErrorMessage

End Sub


Private Sub ClearListview()
On Error GoTo Err

 Dim j, i As Integer
       
                j = ListView1.ListItems.Count
            
        For i = 1 To j
                ListView1.ListItems(i).Checked = False
        Next i
Exit Sub
Err:
ErrorMessage
    
End Sub

Private Sub cmdCancel_Click()
        Set rsAPPROVERS = New clsODASApprovers
        rsAPPROVERS.cancelRECORD
        Set rsAPPROVERS = Nothing
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err

        Set rsAPPROVERS = New clsODASApprovers
        rsAPPROVERS.DeleteRecord
        Set rsAPPROVERS = Nothing

Exit Sub

Err:
    ErrorMessage
End Sub







Private Sub Form_Activate()
'        enableButtons
        disableALLRECORD
        Set rsAPPROVERS = New clsODASApprovers
        rsAPPROVERS.DisplayOperationType
        Set rsAPPROVERS = Nothing
        GetUserCode
        GetAPPROVERS

End Sub


Private Sub Form_Unload(Cancel As Integer)
        baddRECORD = False
        bsearchRECORD = False
        beditRECORD = False
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo Err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
            
            Set rsAPPROVERS = New clsODASApprovers
            rsAPPROVERS.pclearRECORD
            Set rsAPPROVERS = Nothing
            
            frmODASPApprovers.txtUserCode.Text = Item.Text
            loadNAMES
            loadUSER
        Else
            Item.Checked = False
        End If
        

Exit Sub

Err:
    ErrorMessage
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo Err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPApprovers.txtUserCode.Text = Item.Text
            Set rsAPPROVERS = New clsODASApprovers
            rsAPPROVERS.SearchRecord
            Set rsAPPROVERS = Nothing
        Else
            Item.Checked = False
        End If
        

Exit Sub

Err:
    ErrorMessage
End Sub

Private Sub loadDETAILS()
On Error GoTo Err

        Set rsCONTROL = New ADODB.Recordset

        strSQL = "Select * from AdminUserRegister where StaffIDNO = '" & frmODASPApprovers.txtUserCode & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .BOF Or .EOF Then
                        frmODASPApprovers.txtDateAssigned.Text = Date
                        frmODASPApprovers.txtLimitAmount.Text = 0
                        frmODASPApprovers.txtNames.Text = !UserName
                        frmODASPApprovers.txtUserCode.Text = !staffidno
                        frmODASPApprovers.cboStatus = "INFORCE"

                Else
                        frmODASPApprovers.txtDateAssigned.Text = !DateAssigned & ""
                        frmODASPApprovers.txtLimitAmount.Text = !LimitAmount & ""
                        frmODASPApprovers.txtNames.Text = !UserName
                        frmODASPApprovers.txtUserCode.Text = !staffidno
                        frmODASPApprovers.cboStatus = !Status & ""
                End If
        End With
        
        rsCONTROL.Close
        strSQL = ""

Exit Sub

Err:
    ErrorMessage
End Sub

Private Sub loadUSER()
On Error GoTo Err

        Set rsCONTROL = New ADODB.Recordset

        strSQL = "Select * from ODASPApprovers, AdminUserRegister where ODASPApprovers.StaffID = '" & frmODASPApprovers.txtUserCode & "' and ADminUserRegister.StaffIDno = ODASPApprovers.StaffID;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .BOF Or .EOF Then
                        frmODASPApprovers.txtDateAssigned.Text = Date
                        frmODASPApprovers.txtLimitAmount.Text = 0
                        frmODASPApprovers.cboStatus = "INFORCE"

                Else
                        frmODASPApprovers.txtDateAssigned.Text = !DateAssigned & ""
                        frmODASPApprovers.txtLimitAmount.Text = !LimitAmount & ""
                        frmODASPApprovers.txtUserCode.Text = !staffidno
                        frmODASPApprovers.cboStatus = !Status & ""
                End If
        End With
        
        rsCONTROL.Close
        strSQL = ""

Exit Sub

Err:
    ErrorMessage
End Sub

Private Sub loadNAMES()
On Error GoTo Err

        Set rsCONTROL = New ADODB.Recordset

        strSQL = "Select * from AdminUserRegister where StaffIDNo = '" & frmODASPApprovers.txtUserCode & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .BOF Or .EOF Then Exit Sub
                frmODASPApprovers.txtNames.Text = !AllNames
        End With
        
        rsCONTROL.Close
        strSQL = ""

Exit Sub

Err:
    ErrorMessage
End Sub

Private Sub loadRECORD()
On Error GoTo Err

        Set rsCONTROL = New ADODB.Recordset

        strSQL = "Select * from ODASPApprovers where StaffIDNO = '" & frmODASPApprovers.txtUserCode & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .BOF Or .EOF Then Exit Sub
                frmODASPApprovers.txtDateAssigned.Text = Date
                frmODASPApprovers.txtLimitAmount.Text = 0
                frmODASPApprovers.txtNames.Text = !UserName
                frmODASPApprovers.txtUserCode.Text = !staffidno
                frmODASPApprovers.cboStatus = "A"
        End With
        
        rsCONTROL.Close
        strSQL = ""

Exit Sub

Err:
    ErrorMessage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "New &Record "
                            baddRECORD = True
                            enableALLRECORD
'                            disableButtons
                            Button.Caption = "&Save Record"
                    Case "&Save Record"
                            Set rsAPPROVERS = New clsODASApprovers
                            rsAPPROVERS.updateRECORD
                            Set rsAPPROVERS = Nothing
                            Button.Caption = "New &Record "
                End Select
            Case "E"
                    editMYRECORD
    
            Case "S"
                    searchMyRecord

            Case "R"
            Case "H"
        End Select
    End With
Exit Sub
Err:
ErrorMessage
End Sub

Private Sub txtDateAssigned_Click()
    frmODASPApprovers.txtDateAssigned.Text = Date
    frmODASPApprovers.txtDateRetired.Text = DateAdd("d", 20, frmODASPApprovers.txtDateAssigned.Text)
End Sub

Private Sub txtReenterPassword_LostFocus()
        With frmODASPApprovers
            If Trim(.txtPassword.Text) <> Trim(.txtReenterPassword.Text) Then
                MsgBox "The Password you entered do not Match", vbOKOnly
            End If
        End With
End Sub

