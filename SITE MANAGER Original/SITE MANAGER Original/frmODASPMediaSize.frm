VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPMediaSize 
   Caption         =   "Media Sizes"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   9225
   Icon            =   "frmODASPMediaSize.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPMediaSize.frx":0442
   ScaleHeight     =   5190
   ScaleWidth      =   9225
   Begin VB.Frame Frame12 
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox cboUnitCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   6000
         TabIndex        =   20
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
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
         Left            =   3600
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtLength 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtMediaSize 
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   6855
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
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1335
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
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   5535
      End
      Begin VB.Frame Frame1 
         Caption         =   "List of All Sizes Per Media"
         Height          =   3015
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   7575
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
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
         Left            =   7800
         TabIndex        =   8
         Top             =   1800
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
            Picture         =   "frmODASPMediaSize.frx":0784
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
            Picture         =   "frmODASPMediaSize.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   4
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
            TabIndex        =   13
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
            Picture         =   "frmODASPMediaSize.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   12
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
            Picture         =   "frmODASPMediaSize.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   11
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
            Picture         =   "frmODASPMediaSize.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   10
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
            Picture         =   "frmODASPMediaSize.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2520
            Width           =   855
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Unit Measure"
         Height          =   255
         Left            =   4800
         TabIndex        =   21
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Width"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Length"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Size "
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1395
         Width           =   1335
      End
      Begin VB.Label lblRelationshipCode 
         Caption         =   "Media Code"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   390
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmODASPMediaSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub loadRECORD()
'On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPMedia,ODASPMediaSize  Where ODASPMediaSize.MediaCode = ODASPMedia.MediaCode and ODASPMedia.MediaCode = '" & frmODASPMediaSize.txtMediaCode.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
        
        If .EOF Or .EOF Then Exit Sub
        
        frmODASPMediaSize.txtMediaCode = !MediaCode
        frmODASPMediaSize.txtMediaDescription = !MediaDescription
        frmODASPMediaSize.txtMediaSize = !MediaSize
        frmODASPMediaSize.txtLength = !Length
        frmODASPMediaSize.txtWidth = !Width
        frmODASPMediaSize.cboUnitCode.Text = UnitCode & ""

    End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cboUnitCode_GotFocus()
        SelectUnitGotFocus
End Sub

Private Sub cboUnitcode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboUnitCode_LostFocus()
        selectUnitLostFocus
End Sub

Private Sub cmdAddNew_Click()
        baddRECORD = True
        clearRECORD
        enableALLRECORD
        disableButtons
End Sub


Private Sub cmdCancel_Click()
        enableButtons
        clearRELN
        disableRELN
        baddRECORD = False
End Sub
Private Sub clearRECORD()
'On Error GoTo err
        With frmODASPMediaSize
            .txtMediaSize.Text = Empty
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdDelete_Click()
'On Error GoTo err

If txtMediaCode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            Set rsCONTROL = New ADODB.Recordset
    
            strSQL = "Select * from ODASPMediaCode Where MediaCode = '" & frmODASPMediaSize.txtMediaCode.Text & "'"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            
            With rsCONTROL
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearALLRECORD
                getALLOPERATIONS
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
        editMYRECORD
End Sub

Private Sub ValidateRECORD()
'On Error GoTo err
        With frmODASPMediaSize
        
                bSaveRECORD = False
                
                If .txtMediaCode.Text = Empty Then
                        MsgBox "The Operation Type MUST be Entered"
                        .txtMediaCode.SetFocus
                
                ElseIf .txtMediaDescription.Text <= Empty Then
                        MsgBox "The Description of the Media cannot be Left Blank"
                        .txtMediaDescription.SetFocus
                
                ElseIf .txtMediaSize.Text = Empty Then
                        MsgBox "The Media size is important Hence cannot be Left Blank"
                        .txtMediaSize.SetFocus
                
                ElseIf .txtLength.Text = Empty Then
                        MsgBox "The Media Length is Required and Cannot be Left Blank"
                        .txtLength.SetFocus
                        
                ElseIf .txtWidth.Text = Empty Then
                        MsgBox "The Width Cannot Be Left blank"
                        .txtWidth.SetFocus
                        
                ElseIf .cboUnitCode.Text = Empty Then
                        MsgBox "The Unit Measure Cannot Be Left Blank"
                        .cboUnitCode.SetFocus
                Else
                        bSaveRECORD = True
                End If

        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveRECORD()
''On Error GoTo err
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASPMediaSize Where MediaCode = '" & frmODASPMediaSize.txtMediaCode.Text & "' and MediaSize = '" & frmODASPMediaSize.txtMediaSize.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
   With rsCONTROL
        If .BOF Or .EOF Then
                .AddNew
                !MediaCode = frmODASPMediaSize.txtMediaCode
                !MediaSize = frmODASPMediaSize.txtMediaSize

                !PreparedBy = CurrentUserName
                !DatePrepared = Date
        End If
        
        !Length = frmODASPMediaSize.txtLength.Text
        !Width = frmODASPMediaSize.txtWidth.Text
        !UnitCode = frmODASPMediaSize.cboUnitCode.Text
        bSaveRECORD = False
        
         .Update
         .Requery
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
        bSaveRECORD = False
        ValidateRECORD
        If bSaveRECORD = True Then
            SaveRECORD
                If bSaveRECORD = False Then
                    enableButtons
                    disableALLRECORD
                    baddRECORD = False
                End If
        End If
        showALLMEDIASIZES

        
End Sub

Private Sub cmdSearch_Click()
        searchMyRecord
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableButtons
    showALLMEDIASIZES
End Sub

Private Sub Form_Load()

    OpenODBCConnection
      
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
        
        If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
                        
            frmODASPMediaSize.txtMediaCode.Text = Item.Text
            loadRECORD
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub



Private Sub txtMediaSize_Click()
On Error GoTo err
        With frmODASPMediaSize
            If .txtLength.Text = Empty Or .txtWidth.Text = Empty Then Exit Sub
            .txtMediaSize.Text = Trim(.txtLength.Text) + " BY " + Trim(.txtWidth.Text) + " " + Trim(.cboUnitCode.Text)
        End With
Exit Sub
err:
    ErrorMessage
End Sub

