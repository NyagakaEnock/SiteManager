VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPLandRate 
   Caption         =   "Land Rate"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10650
   Icon            =   "frmODASPLandRate.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPLandRate.frx":0442
   ScaleHeight     =   7050
   ScaleWidth      =   10650
   Begin VB.Frame Frame12 
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10455
      Begin VB.Frame Frame2 
         Height          =   3735
         Left            =   5400
         TabIndex        =   11
         Top             =   120
         Width           =   4935
         Begin VB.TextBox txtRateSQR 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1680
            TabIndex        =   35
            Top             =   3240
            Width           =   1095
         End
         Begin VB.TextBox txtMediaCode 
            BackColor       =   &H00FFFFC0&
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
            Left            =   1680
            TabIndex        =   12
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtWidth 
            Height          =   285
            Left            =   3720
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtLandRate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Left            =   3720
            TabIndex        =   24
            Top             =   3240
            Width           =   1095
         End
         Begin VB.TextBox txtPaymentModeDescription 
            BackColor       =   &H00FFFFC0&
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
            Left            =   1680
            TabIndex        =   23
            Top             =   2808
            Width           =   3135
         End
         Begin VB.ComboBox cboPaymentMode 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1680
            TabIndex        =   21
            Top             =   2380
            Width           =   3135
         End
         Begin VB.TextBox txtTown 
            BackColor       =   &H00FFFFC0&
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
            Left            =   1680
            TabIndex        =   20
            Top             =   1952
            Width           =   3135
         End
         Begin VB.TextBox txtTownCode 
            BackColor       =   &H00FFFFC0&
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
            Left            =   1680
            TabIndex        =   18
            Top             =   1524
            Width           =   3135
         End
         Begin VB.TextBox txtMediaSize 
            BackColor       =   &H00FFFFC0&
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
            Left            =   1680
            TabIndex        =   17
            Top             =   1096
            Width           =   3135
         End
         Begin VB.TextBox txtMediaDescription 
            BackColor       =   &H00FFFFC0&
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
            Left            =   1680
            TabIndex        =   14
            Top             =   668
            Width           =   3135
         End
         Begin VB.Label Label8 
            Caption         =   "Rate/ SQUARE MTR"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Town"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1982
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Mode Description"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Rate Payable"
            Height          =   255
            Left            =   2760
            TabIndex        =   25
            Top             =   3270
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Payment Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2410
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Town Code"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1554
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Size "
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1126
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Media Description"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   726
            Width           =   1335
         End
         Begin VB.Label lblRelationshipCode 
            Caption         =   "Media Code"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.TextBox txtLength 
         Height          =   285
         Left            =   9120
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Media Sizes"
         Height          =   1815
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   5175
         Begin MSComctlLib.ListView ListView1 
            Height          =   1455
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   2566
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
         Caption         =   "Media Codes"
         Height          =   1935
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   5175
         Begin MSComctlLib.ListView ListView2 
            Height          =   1575
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   2778
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
      Begin VB.Frame Frame1 
         Caption         =   "List of All Land Rates Per Town"
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   9015
         Begin MSComctlLib.ListView ListView3 
            Height          =   2535
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   9240
         TabIndex        =   3
         Top             =   3840
         Width           =   1095
         Begin VB.CommandButton cmdAddNew 
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
            Picture         =   "frmODASPLandRate.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
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
            Picture         =   "frmODASPLandRate.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   990
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
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
            Picture         =   "frmODASPLandRate.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1365
            Width           =   855
         End
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
            Picture         =   "frmODASPLandRate.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1740
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
            Picture         =   "frmODASPLandRate.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2115
            Width           =   855
         End
         Begin VB.CommandButton cmdPrint 
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
            Picture         =   "frmODASPLandRate.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2520
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmODASPLandRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub loadRECORD()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPMedia,ODASPMediaSize  Where ODASPMediaSize.MediaCode = ODASPMedia.MediaCode and ODASPMedia.MediaCode = '" & frmODASPLandRate.txtMediaCode.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
        
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPLandRate.txtMediaSize = !MediaSize

    End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadMEDIA()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPMedia  "
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPLandRate.txtMediaCode = !MediaCode
        frmODASPLandRate.txtMediaDescription = !MediaDescription
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cboPaymentMode_GotFocus()
        selectPaymentModeGotFocus
End Sub

Private Sub cboPaymentMode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboPaymentMode_LostFocus()
        selectPaymentModeLostFocus
End Sub


Private Sub cmdAddNew_Click()
        baddRECORD = True
        clearRECORD
        enableALLRECORD
        disableButtons
        Me.txtLandRate.Locked = True
End Sub

Private Sub cmdCancel_Click()
        enableButtons
        clearALLRECORD
        disableALLRECORD
        baddRECORD = False
End Sub
Private Sub clearRECORD()
On Error GoTo err
        With frmODASPLandRate
            .txtMediaSize.Text = Empty
            .txtLength.Text = ""
            .txtWidth.Text = ""
'            .txtRateSQR.Text = ""
            .txtLandRate.Text = ""
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdDelete_Click()
On Error GoTo err

If txtMediaCode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            Set rsCONTROL = New ADODB.Recordset
    
            strSQL = "Select * from ODASPMediaCode Where MediaCode = '" & frmODASPLandRate.txtMediaCode.Text & "'"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            
            With rsCONTROL
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearALLRECORD
                'getALLOPERATIONS
            End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cmdEdit_Click()
        EditMyRecord
End Sub

Private Sub validateRECORD()
On Error GoTo err
        With frmODASPLandRate
        
                bsaveRECORD = False
                
                If .txtMediaCode.Text = Empty Then
                        MsgBox "The Operation Type MUST be Entered"
                        .txtMediaCode.SetFocus
                
                ElseIf .txtMediaDescription.Text <= Empty Then
                        MsgBox "The Description of the Media cannot be Left Blank"
                        .txtMediaDescription.SetFocus
                
                ElseIf .txtMediaSize.Text = Empty Then
                        MsgBox "The Media size is important Hence cannot be Left Blank"
                        .txtMediaSize.SetFocus
                
                ElseIf .txtTownCode.Text = Empty Then
                        MsgBox "The Media TownCode is Required and Cannot be Left Blank"
                        .txtTownCode.SetFocus
                        
                ElseIf .txtRateSQR.Text = Empty Then
                        MsgBox "The Land Rate Per SQURE MTR Cannot Be Left blank"
                        .txtRateSQR.SetFocus
                        
                ElseIf .cboPaymentMode.Text = Empty Then
                        MsgBox "The Unit Measure Cannot Be Left Blank"
                        .cboPaymentMode.SetFocus
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
    With frmODASPLandRate
    
            Set rsFindRecord = New ADODB.Recordset
            strSQL = "Select * from ODASPLandRate Where MediaCode = '" & .txtMediaCode.Text & "' and MediaSize = '" & .txtMediaSize.Text & "' and PaymentMode = '" & .cboPaymentMode.Text & "' and townCode = '" & .txtTownCode & "' "
            rsFindRecord.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsFindRecord.BOF Or rsFindRecord.EOF Then
                    rsFindRecord.AddNew
                    rsFindRecord!MediaCode = .txtMediaCode
                    rsFindRecord!MediaSize = .txtMediaSize
                    rsFindRecord!townCode = .txtTownCode.Text
                    rsFindRecord!PaymentMode = .cboPaymentMode.Text
                    rsFindRecord!CurrentYear = Year(Date)
                    rsFindRecord!Preparedby = CurrentUserName
                    rsFindRecord!dateprepared = Date
            End If
                
            rsFindRecord!Amount = CDbl(.txtLandRate.Text)
            rsFindRecord!PerSQRMtr = CDbl(.txtRateSQR.Text)
            bsaveRECORD = False
                
            rsFindRecord.Update
    End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
'            rsCONTROL.CancelUpdate
'            rsCONTROL.Requery
    Else
        UpdateErrorMessage
    End If

End Sub
Private Sub cmdUpdate_Click()
        bsaveRECORD = False

        validateRECORD

        Me.txtLandRate.Text = Me.txtLength.Text * Me.txtWidth.Text * Me.txtRateSQR.Text
        If bsaveRECORD = True Then
            saveRecord
                If bsaveRECORD = False Then
                    enableButtons
                    disableALLRECORD
                    baddRECORD = False
                End If
        End If
        showALLLANDRATES

        
End Sub

Private Sub cmdSearch_Click()
        searchMyRecord
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableButtons
    showALLMEDIA2
'    showALLMEDIASIZES
    showALLLANDRATES
End Sub

Private Sub Form_Load()

    OpenConnection
      
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
                        
            frmODASPLandRate.txtMediaSize.Text = Item.SubItems(2)
            frmODASPLandRate.txtMediaDescription.Text = Item.SubItems(1)
            frmODASPLandRate.txtLength.Text = Item.SubItems(3)
            frmODASPLandRate.txtWidth.Text = Item.SubItems(4)
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPLandRate.txtMediaCode.Text = Item.Text
'            loadMEDIA
            showALLMEDIASIZES
            frmODASPLandRate.txtRateSQR.Text = ""
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub txtMediaSize_Click()
On Error GoTo err
        With frmODASPLandRate
            If .txtTownCode.Text = Empty Or .txtLandRate.Text = Empty Then Exit Sub
            .txtMediaSize.Text = Trim(.txtTownCode.Text) + " BY " + Trim(.txtLandRate.Text) + " " + Trim(.cboPaymentMode.Text)
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub txtRateSQR_LostFocus()
With Me
    If .txtLength.Text = Empty Then Exit Sub
    .txtLandRate.Text = .txtLength.Text * .txtWidth.Text * .txtRateSQR.Text
End With

End Sub
