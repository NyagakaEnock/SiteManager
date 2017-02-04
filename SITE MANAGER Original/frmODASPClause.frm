VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASPClause 
   Caption         =   "EDITING THE CONTRACT CLAUSES"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11895
      Begin VB.TextBox txtClause 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Text            =   " "
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox cboClauseCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         ItemData        =   "frmODASPClause.frx":0000
         Left            =   240
         List            =   "frmODASPClause.frx":0002
         TabIndex        =   6
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtClauseDescription 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Text            =   " "
         Top             =   240
         Width           =   7455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Clause Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   11655
         Begin VB.TextBox txtClauseDetails 
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
            Height          =   3135
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Text            =   "frmODASPClause.frx":0004
            Top             =   240
            Width           =   11415
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Clause No"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   1164
      ButtonWidth     =   3069
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
               Picture         =   "frmODASPClause.frx":0006
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":0680
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":0BC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":1014
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":132E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":19A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":2022
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":2474
               Key             =   ""
            EndProperty
         EndProperty
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
               Picture         =   "frmODASPClause.frx":2AEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":3168
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":35BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":38D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":3F4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":45C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASPClause.frx":4A1A
               Key             =   ""
            EndProperty
         EndProperty
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
End
Attribute VB_Name = "frmODASPClause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboClauseCode_GotFocus()
 With Me
        If .cboClauseCode.ListCount <> 0 Then Exit Sub
            .cboClauseCode.Clear
            AttachSQL = "SELECT (ODASPClause.ClauseDescription)as selectfield,ODASPClause.* FROM ODASPClause ORDER BY Clause;"
            AttachDropDowns
    End With
End Sub

Private Sub cboClauseCode_LostFocus()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASPClause WHERE ClauseDescription = '" & .txtClause.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
            .txtClauseDescription.Text = rsFindRecord!ClauseDescription
            .txtClause.Text = rsFindRecord!Clause
            If Not IsNull(rsFindRecord!ClauseDetails) Then
               .txtClauseDetails.Text = rsFindRecord!ClauseDetails
            Else:
              MsgBox "No Clause Details Defined For This Clause,Please Add"
              .txtClauseDetails.Text = Empty
            End If
        
    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub Form_Load()
With Me
End With
End Sub
Private Sub ValidateRECORD()
On Error GoTo err
        With frmODASPClause
                
                If .txtClause.Text <= " " Then
                    MsgBox " You MUST Select The Clause No. You Want To Add!!"
                    .txtClause.SetFocus
                ElseIf .txtClauseDetails.Text <= " " Then
                    MsgBox "Enter The Clause Details For This Clause Number Under This Contract!!"
                    .txtClauseDetails.SetFocus
               
                Else
                        bSaveRECORD = True
                End If
                
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub saveClause()
On Error GoTo err

    With frmODASPClause
    

            Set rsSAVE = New ADODB.Recordset
            
            strSQL = "Select * from ODASPClause Where Clause = '" & .txtClause.Text & "'"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.BOF Or rsSAVE.EOF Then
                 rsSAVE.AddNew
                rsSAVE!Clause = .txtClause.Text
                rsSAVE!ClauseDescription = .txtClauseDescription.Text
                rsSAVE!PreparedBY = CurrentUserName
                rsSAVE!DatePrepared = Date
            End If
                rsSAVE!ClauseDetails = .txtClauseDetails.Text
            
            bSaveRECORD = False

            rsSAVE.Update
            rsSAVE.Requery
  End With

Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If

End Sub
Public Sub DeleteClause()
On Error GoTo err

    With frmODASPClause

            Set rsSAVE = New ADODB.Recordset
            
            strSQL = "Select * from ODASPClause Where Clause = '" & .cboClauseCode.Text & "'"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                  rsSAVE!ClauseDetails = Empty
            
            bSaveRECORD = False

            rsSAVE.Update
            rsSAVE.Requery
  End With

Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "New &Record"
                        baddRECORD = True
                        enableALLRECORD
                        Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Save Record"
                            bSaveRECORD = True
                            ValidateRECORD
                            If bSaveRECORD = True Then
                                saveClause
                                    If bSaveRECORD = False Then
                                        
                                        disableALLRECORD
                                       
                                    End If
                            
                           baddRECORD = False: Button.Caption = "Next &Record": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                        End If
                   Case "Next &Record"
                        enableALLRECORD
                        Me.cboClauseCode.Text = Empty
                        Me.txtClauseDescription.Text = Empty
                        Me.txtClauseDetails.Text = Empty
                       Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                   End Select
            Case "E"
                Select Case Button.Caption
                    Case "&Edit/Change "
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                
                    Case "Save &Changes "
                    
                        saveClause
                        beditRECORD = False: Button.Caption = "&Edit/Change ": Button.Image = 6
                    End Select
                    
            Case "P"
                 
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                        bsearchRECORD = True
                        Button.Caption = "Delete"
                    Case "Cancel"
                        cancelCMD
                    Case "Delete"
                           DeleteClause
                          DeleteClause
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
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen

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

Private Sub txtClause_LostFocus()
With Me
 cboClauseCode_LostFocus
End With
End Sub
